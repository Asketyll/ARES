' Module: PropertyRendering
' Description: The SOLE text writer of ARES. Displays a custom property's value inside a text by fully
'              replacing a "Prop[Name]" token; the binding lives in hidden ARES_SYS metadata, never in
'              the visible text or the graphic group. Full mechanism (binding storage, the 4-branch
'              release state machine, value semantics): see _bmad/docs/property-rendering-mechanics.md.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (global ARESConfig), CustomPropertyHandler, PropertyTagging,
'               PropertyCalculation, StringsInEl, Link, LangManager, ErrorHandlerClass (global ErrorHandler),
'               CallStackClass (global CallStack)
'
' Never writes a USER property value or attaches/detaches anything itself (delegated to PropertyTagging).
' Any SetPropertyValueToElement on the "ARES" library inside this module is a review BLOCKER.

Option Explicit

' Token grammar - deliberately the exact form the calc grammar already uses on its left-hand side
' (Prop[name]=Source), so the user learns ONE spelling. The "]" terminator makes prefix collisions
' impossible, so no longest-match rule is needed; the tag/calc grammar already forbids "[", "]" and ";"
' inside a property name.
Private Const TOKEN_OPEN As String = "Prop["
Private Const TOKEN_CLOSE As String = "]"

' The two String properties of the ARES_Render ItemType. Every access names BOTH of them and the library
' explicitly - see ReadRenderMetadata for why an omitted ItemName silently breaks every read and write.
Private Const PROP_SCHEMA As String = "SchemaVersion"
Private Const PROP_ENTRIES As String = "Entries"

' One parsed ARES_Render entry: which sub-text it drives, the Template, and the value last rendered for
' each of the Template's tokens. ValNames/ValValues are parallel, bounded by nVals. Dropped marks an
' entry the state machine released; serialisation skips it, which avoids ever copying a UDT holding
' dynamic arrays just to compact the list.
Private Type RenderEntry
    SubId As Long
    Template As String
    ValNames() As String
    ValValues() As String
    nVals As Long
    Dropped As Boolean
End Type

' Outcomes of the per-entry state machine.
Private Const ENTRY_UNCHANGED As Long = 0
Private Const ENTRY_UPDATED As Long = 1
Private Const ENTRY_DROP As Long = 2

' One-shot status guards - each refusal surfaces once per PROCESSED ELEMENT, reset in ProcessElement.
' Same shape as PropertyCalculation's mbRejectedShown / mbNoTargetShown / mbMultiShown.
Private mbTokenUnknownShown As Boolean
Private mbValueUnsupportedShown As Boolean
Private mbValueIllegalShown As Boolean
Private mbMetadataInvalidShown As Boolean
Private mbMetadataUnreadableShown As Boolean
Private mbSchemaShown As Boolean
Private mbLibraryMissingShown As Boolean
Private mbAmbiguousShown As Boolean
Private mbBindingReleasedShown As Boolean

Private mbLockedShown As Boolean
Private mbDriftShown As Boolean
Private mbDuplicateShown As Boolean
Private mbAdjacentShown As Boolean
Private mbTextNodeRefusedShown As Boolean
Private mbNotBoundShown As Boolean
Private mbGovernedShown As Boolean
Private mbCycleShown As Boolean

' D6 - WHITELIST (not a blacklist) of symbols an addition may border a value with: a blacklist could
' silently miss a forging character (e.g. a locale thousands separator) - this list is closed by
' construction instead. Full admission criterion and named exceptions: see
' property-rendering-mechanics.md.
Private Const SAFE_BOUNDARY_SYMBOLS As String = "%()[]{}\*#~<>=:;!?&@_|"""

' D8 - characters that can appear INSIDE a number literal (any base), bounding NumericTailIsPossible's
' scan only - never granting safety by itself (too broad is harmless, too narrow just stops the scan
' early). Full rationale: see property-rendering-mechanics.md.
Private Const NUMBER_CAPABLE_CHARS As String = "0123456789abcdefABCDEFxXbBoO"

' The two self-disable conditions also write an English log line. Those flags are SESSION-scoped, not
' per-element: ResetOneShots must not clear them, or a station whose DGNLib predates epic 15 would log
' one line per processed element for the whole session.
Private mbLibraryLogged As Boolean
Private mbSchemaLogged As Boolean

' Third SESSION-scoped self-disable: guards against an unbounded write/restore fault loop, not a normal
' refusal. Full rationale: see "Session-scoped self-disable guards" in property-rendering-mechanics.md.
Private mbWriteDisabled As Boolean

' Managed property-name cache. GetCustomPropertyNames() has NO cache of its own: every call builds a
' New ItemTypeLibraries, does FindForDesignFile + Refresh, then walks Find("*") with ReDim Preserve.
' Validating one token per element per pass through that would be ruinous. Invalidated when the active
' design file changes (the DGN-open invalidation point, observed without needing a new call site).
Private msPropNames() As String
Private mnPropNames As Long
Private mbNamesCached As Boolean
Private msCachedFor As String

' ARES_SYS presence, cached alongside the names and invalidated by the same file change. Absent library
' = self-disable, fail-closed (no bind, no render, never a partial write).
Private mbSysChecked As Boolean
Private mbSysPresent As Boolean

' Bounded repaint hop. PropertyCalculation appends the elements whose value it just changed; the drain
' runs the FULL state machine on each of them AND on their graphic-group siblings, so a sibling the
' pipeline does not re-queue still refreshes inside the same batch.
Private moDirty() As element
Private mnDirty As Long
Private mbDraining As Boolean

'######################################################################################################################
'                                          PUBLIC SURFACE
'######################################################################################################################

' Master switch (ARES_Text_Render, OFF by default). Mirrors PropertyCalculation.IsEnabled: lazy config
' init, fail-closed False on any nil.
Public Function IsEnabled() As Boolean
    On Error GoTo ErrorHandler

    IsEnabled = False
    If ARESConfig Is Nothing Then Exit Function
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    If ARESConfig.ARES_TEXT_RENDER Is Nothing Then Exit Function
    IsEnabled = CBool(ARESConfig.ARES_TEXT_RENDER.Value)
    Exit Function

ErrorHandler:
    IsEnabled = False
End Function

' Drop the cached property names and the ARES_SYS presence flag, and clear the session-scoped self-disable
' state. Called by the test harness; the runtime invalidates itself when the active design file changes.
' It only assigns module variables and so cannot realistically fault, but it is Public - it therefore
' carries the standard handler rather than an exception the project's blocker list would have to carve out.
Public Sub RefreshRenderCaches()
    On Error GoTo ErrorHandler

    mbNamesCached = False
    mnPropNames = 0
    msCachedFor = ""
    mbSysChecked = False
    mbSysPresent = False
    mbLibraryLogged = False
    mbSchemaLogged = False
    mbWriteDisabled = False
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.RefreshRenderCaches"
End Sub

' The Depth-0 hook, called from ElementChangeHandler.ProcessElement right after Calculation and before the
' graphic-group filter. Guard order and the v1 queue-only limitation: see "ProcessElement hook" in
' property-rendering-mechanics.md.
Public Sub ProcessElement(ByVal oEl As element)
    On Error GoTo ErrorHandler
    Dim bStackPushed As Boolean

    ResetOneShots

    If Not IsEnabled Then Exit Sub
    If oEl Is Nothing Then Exit Sub

    ' Pushed only past the cheap enable/validity guards above. Every guard below this point funnels
    ' through the single DrainAndExit label, so ONE Pop there (plus one in ErrorHandler) covers all
    ' of them - no per-guard Pop needed, unlike modules without a shared exit label.
    CallStack.Push "PropertyRendering.ProcessElement", oEl
    bStackPushed = True

    ' A metadata write already failed this session: stay inert instead of replaying the same failing
    ' write, and its restore, on every single pass. The list still drains so nothing leaks across batches.
    If mbWriteDisabled Then GoTo DrainAndExit

    ' Cheap TYPE filter FIRST: Text / TextNode / Cell only. Dimensions, notes, tables and tag elements are
    ' out of v1 scope - their tokens stay literal. This runs on the handle as received, deliberately: it
    ' reads only IsTextElement / IsTextNodeElement / IsCellElement, and an element's TYPE cannot go stale.
    ' Keeping it ahead of the re-fetch is what stops every line and arc in the queue paying a COM call.
    If Not IsRenderableType(oEl) Then GoTo DrainAndExit

    ' RE-FETCH BEFORE READING ANY TEXT: the handle handed to this hook is not necessarily current (the
    ' idle batch materialises all elements up front). See "The stale-handle invariant (FreshHandle)" in
    ' property-rendering-mechanics.md.
    Set oEl = FreshHandle(oEl)
    If oEl Is Nothing Then GoTo DrainAndExit

    ' IsLocked comes AFTER the re-fetch - unlike the type, a lock CAN have been taken since the handle was
    ' captured - and BEFORE any state decision: a locked element must never produce a transition.
    If oEl.IsLocked Then
        ReportLocked
        GoTo DrainAndExit
    End If

    Dim ids() As Long
    Dim texts() As String
    Dim nBearers As Long

    ' ONE walk feeds both branches (and every entry of the bound branch): the ordinals and the whole text
    ' of each text-bearing sub-element. -1 = the walk faulted, so no ordinal can be trusted.
    nBearers = StringsInEl.EnumerateTextSubIds(oEl, ids, texts)
    If nBearers <= 0 Then GoTo DrainAndExit

    ' Self-disable fail-closed when the internal library is missing (old resource, new .mvba).
    If Not IsSysLibraryPresent() Then
        ReportLibraryMissing
        GoTo DrainAndExit
    End If

    If CustomPropertyHandler.IsItemAttachedToElement(oEl, ARES_ITEM_RENDER, ARES_NAME_LIBRARY_SYS) Then
        RenderBoundElement oEl, texts, nBearers
    Else
        TryFirstAuthor oEl, texts, nBearers, False
    End If

DrainAndExit:
    DrainRepaintHop
    If bStackPushed Then CallStack.Pop
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.ProcessElement"
    ' The hop must never leak across batches, so it drains even on a fault (the drain is re-entrance
    ' guarded and clears the list unconditionally).
    DrainRepaintHop
    If bStackPushed Then CallStack.Pop
End Sub

' "Does this engine own that text?" - True when ARES_Render is attached. IsEnabled tested FIRST so a
' render-free configuration pays no COM cost.
' NO PRODUCTION CALLER since the Auto Lengths removal - kept as a meaningful state query and a seam
' PropertyRenderingTest asserts against. Do not delete it as unused.
Public Function IsRenderBound(ByVal El As element) As Boolean
    On Error GoTo ErrorHandler

    IsRenderBound = False
    If Not IsEnabled Then Exit Function
    If El Is Nothing Then Exit Function

    IsRenderBound = CustomPropertyHandler.IsItemAttachedToElement(El, ARES_ITEM_RENDER, ARES_NAME_LIBRARY_SYS)
    Exit Function

ErrorHandler:
    IsRenderBound = False
End Function

' Containment seam consumed by PropertyCalculation's CellText source: the SubIds of oCell's sub-texts this
' engine writes, so a rendered value can never feed the CellText[...] value that governs it. Keyed on
' metadata PRESENCE only, never IsEnabled - see "Containment (GetExcludedSubIds)" in
' property-rendering-mechanics.md.
Public Function GetExcludedSubIds(ByVal oCell As element, ByRef ids() As Long, ByRef nIds As Long) As Boolean
    On Error GoTo ErrorHandler

    GetExcludedSubIds = False
    nIds = 0
    If oCell Is Nothing Then Exit Function

    ' Two cheap, staleness-proof gates before paying for the re-fetch below - calc calls this on every
    ' CellText read. See property-rendering-mechanics.md.
    If Not IsSysLibraryPresent() Then Exit Function
    If Not CustomPropertyHandler.IsItemAttachedToElement(oCell, ARES_ITEM_RENDER, ARES_NAME_LIBRARY_SYS) Then Exit Function

    ' Now re-fetch, for the same reason as ProcessElement: this seam runs a TEXT walk to resolve each
    ' entry's SubId, calc calls it mid-batch, and a stale text would resolve the exclusion onto the wrong
    ' sub-text - silently changing a CellText value instead of protecting it.
    Set oCell = FreshHandle(oCell)
    If oCell Is Nothing Then Exit Function

    Dim ents() As RenderEntry
    Dim nEnts As Long
    Dim texts() As String
    Dim walkIds() As Long
    Dim subIds() As Long
    Dim nBearers As Long
    Dim i As Long

    ' Attached but unreadable/vandalised: ReadRenderMetadata surfaces the RIGHT status itself (a schema
    ' mismatch has its own, and must not be overwritten by the generic one). The dangerous fallback - the
    ' calc value silently ingesting rendered text - is made visible rather than guessed at.
    If Not ReadRenderMetadata(oCell, ents, nEnts) Then Exit Function
    If nEnts = 0 Then Exit Function

    ' The STORED SubId is NOT the answer (ordinal drift). Resolve exactly the way the renderer does, so the
    ' exclusion set matches what it would actually write. See property-rendering-mechanics.md.
    nBearers = StringsInEl.EnumerateTextSubIds(oCell, walkIds, texts)
    If nBearers <= 0 Then Exit Function
    ResolveAllSubIds ents, nEnts, texts, nBearers, subIds

    ReDim ids(0 To nEnts - 1)
    For i = 0 To nEnts - 1
        ' An entry that does not resolve is one the renderer itself refuses to write this pass, so it has
        ' nothing to hide from the value: skip it rather than exclude a sub-text at random.
        If subIds(i) >= 0 Then
            ids(nIds) = subIds(i)
            nIds = nIds + 1
        End If
    Next i

    If nIds = 0 Then Exit Function
    GetExcludedSubIds = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.GetExcludedSubIds"
    GetExcludedSubIds = False
    nIds = 0
End Function

' Append entry point for the bounded repaint hop, called by PropertyCalculation.ApplyValueToSibling on a
' value STATE CHANGE. Stores the SOURCE ELEMENT, not its group id (Link.GetLink excludes the element it is
' given). Full rationale: see "Repaint hop (bounded)" in property-rendering-mechanics.md.
Public Sub NoteDirtyGroup(ByVal oSource As element)
    On Error GoTo ErrorHandler

    If Not IsEnabled Then Exit Sub
    If oSource Is Nothing Then Exit Sub

    ReDim Preserve moDirty(0 To mnDirty)
    Set moDirty(mnDirty) = oSource
    mnDirty = mnDirty + 1
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.NoteDirtyGroup"
End Sub

' Drain the repaint hop: for each noted element, run the FULL state machine on it and then on its
' graphic-group siblings, ONE Link.GetLink per distinct group. Never a blind re-render. Re-entrance-guarded;
' the list clears unconditionally. Two load-bearing rules (re-fetch before each run; each element runs at
' most once) - see "Repaint hop (bounded)" in property-rendering-mechanics.md.
Public Sub DrainRepaintHop()
    On Error GoTo ErrorHandler

    If mbDraining Then Exit Sub
    If mnDirty = 0 Then Exit Sub
    mbDraining = True

    Dim seen() As Long
    Dim nSeen As Long
    Dim doneIds() As String
    Dim nDone As Long
    Dim els() As element
    Dim src As element
    Dim bWalk As Boolean
    Dim i As Long, j As Long

    nSeen = 0
    nDone = 0
    For i = 0 To mnDirty - 1
        Set src = FreshHandle(moDirty(i))
        If Not src Is Nothing Then
            ' Nested Ifs, never an And chain: .GraphicGroup RAISES on a non-graphical element and VBA
            ' evaluates both operands of And.
            bWalk = False
            If src.IsGraphical Then
                If src.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID Then
                    If Not NoteGroupSeen(src.GraphicGroup, seen, nSeen) Then
                        els = Link.GetLink(src)
                        bWalk = HasElements(els)
                    End If
                End If
            End If

            RenderDrainTarget src, doneIds, nDone

            If bWalk Then
                For j = LBound(els) To UBound(els)
                    RenderDrainTarget els(j), doneIds, nDone
                Next j
            End If
        End If
    Next i

    mnDirty = 0
    Erase moDirty
    mbDraining = False
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.DrainRepaintHop"
    mnDirty = 0
    Erase moDirty
    mbDraining = False
End Sub

' The BindPropertyRender key-in's worker: author-time validation, first bind and first render on ONE
' element. Returns True when the element ends up bound. Also the manual entry point for the case the
' hybrid auto-bind deliberately refuses (a token whose property is not attached yet).
Public Function BindElement(ByVal El As element) As Boolean
    On Error GoTo ErrorHandler

    BindElement = False

    ' The key-in is a fresh user gesture, so it starts from a clean slate: without this, a one-shot flag
    ' already raised by an automatic pass would make the bind refuse in COMPLETE silence, and across a
    ' multi-element selection only the first refusal would ever reach the status bar.
    ResetOneShots

    If El Is Nothing Then Exit Function

    ' Self-disabled after a failed metadata write - a refusal the user asked for, so it must stay visible.
    ' See "First-author key-in visibility (BindElement)" in property-rendering-mechanics.md.
    If mbWriteDisabled Then
        ReportMetadataUnreadable
        Exit Function
    End If

    ' Same reason as ProcessElement: the key-in loops over a selection set materialised before the first
    ' bind, so element k's render can leave element k+n's handle serving a pre-write text.
    Set El = FreshHandle(El)
    If El Is Nothing Then Exit Function

    If Not IsRenderableType(El) Then Exit Function

    If El.IsLocked Then
        ReportLocked
        Exit Function
    End If

    If Not IsSysLibraryPresent() Then
        ReportLibraryMissing
        Exit Function
    End If

    Dim ids() As Long
    Dim texts() As String
    Dim nBearers As Long
    nBearers = StringsInEl.EnumerateTextSubIds(El, ids, texts)
    If nBearers <= 0 Then Exit Function

    ' Already bound: nothing to author, just bring it up to date through the full state machine.
    If CustomPropertyHandler.IsItemAttachedToElement(El, ARES_ITEM_RENDER, ARES_NAME_LIBRARY_SYS) Then
        RenderBoundElement El, texts, nBearers
        BindElement = True
        Exit Function
    End If

    BindElement = TryFirstAuthor(El, texts, nBearers, True)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.BindElement"
    BindElement = False
End Function

'######################################################################################################################
'                                   TEMPLATE MODEL (pure logic - Public test seams)
'######################################################################################################################

' Expand a Template (L0 T0 L1 T1 ... Ln) against a set of values; an EMPTY/ABSENT value renders the token's
' OWN LITERAL TEXT. bOk reports whether the expansion is TRUSTWORTHY (fails OPEN on a fault) - see
' "ExpandTemplate" in property-rendering-mechanics.md.
Public Function ExpandTemplate(ByVal sTemplate As String, ByRef ValNames() As String, ByRef ValValues() As String, ByVal nVals As Long, Optional ByVal bValidateNames As Boolean = True, Optional ByRef bOk As Boolean) As String
    On Error GoTo ErrorHandler

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim i As Long
    Dim sOut As String
    Dim sVal As String

    bOk = False
    ExpandTemplate = sTemplate
    If Not ParseTemplate(sTemplate, lits, toks, nTok, bValidateNames) Then Exit Function
    ' The parse held, so sTemplate IS its own expansion when it carries no token.
    bOk = True
    If nTok = 0 Then Exit Function

    sOut = ""
    For i = 0 To nTok - 1
        sOut = sOut & lits(i)
        sVal = LookupValue(toks(i), ValNames, ValValues, nVals)
        If Len(sVal) = 0 Then
            sOut = sOut & TokenLiteral(toks(i))
        Else
            sOut = sOut & sVal
        End If
    Next i
    sOut = sOut & lits(nTok)

    ExpandTemplate = sOut
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.ExpandTemplate"
    ExpandTemplate = sTemplate
    bOk = False
End Function

' Align a user-edited visible string against the Template that produced the last rendering, deriving the
' new Template plus the surviving LastValues (the per-token release). Full mechanism (literal walk, span
' survival, why anchor-first): see "AlignVisible (per-token release, literal walk)" in
' property-rendering-mechanics.md.
Public Function AlignVisible(ByVal sVisible As String, ByVal sTemplate As String, ByRef ValNames() As String, ByRef ValValues() As String, ByVal nVals As Long, ByRef sNewTemplate As String, ByRef NewNames() As String, ByRef NewValues() As String, ByRef nNew As Long, Optional ByVal bValidateNames As Boolean = True) As Boolean
    On Error GoTo ErrorHandler

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim i As Long
    Dim cursor As Long
    Dim pos As Long
    Dim sSpan As String
    Dim sKeep As String
    Dim sKeepAfter As String
    Dim sLit As String
    Dim sAnchor As String
    Dim sEntry As String
    Dim bHasEntry As Boolean
    Dim bSurvived As Boolean
    Dim bSpanFound As Boolean
    Dim lCmp As Long
    Dim sOut As String

    AlignVisible = False
    sNewTemplate = ""
    nNew = 0

    If Not ParseTemplate(sTemplate, lits, toks, nTok, bValidateNames) Then Exit Function
    If nTok = 0 Then Exit Function

    ' L0 must sit at the very start (empty when the Template opens with a token), matched case-insensitively
    ' like the tokeniser. A surviving span is re-materialised CANONICALLY (TokenLiteral), an edited one keeps
    ' the user's own casing verbatim - see "AlignVisible" in property-rendering-mechanics.md.
    If Len(lits(0)) > 0 Then
        If StrComp(Left(sVisible, Len(lits(0))), lits(0), vbTextCompare) <> 0 Then Exit Function
    End If
    cursor = Len(lits(0)) + 1
    sOut = Left(sVisible, Len(lits(0)))

    For i = 0 To nTok - 1
        bHasEntry = HasValueEntry(toks(i), ValNames, nVals)
        sEntry = LookupValue(toks(i), ValNames, ValValues, nVals)

        ' What the LAST rendering put in this span is KNOWN: the value itself, or - in the unset state -
        ' the token's own literal text.
        sAnchor = ""
        lCmp = vbBinaryCompare
        If bHasEntry Then
            If Len(sEntry) > 0 Then
                sAnchor = sEntry
            Else
                sAnchor = TokenLiteral(toks(i))
                lCmp = vbTextCompare
            End If
        End If

        ' ANCHOR FIRST, literal search only as fallback: locating the closing literal blindly breaks the
        ' moment a VALUE contains it (values are copied verbatim). See property-rendering-mechanics.md.
        bSpanFound = False
        If Len(sAnchor) > 0 Then
            If SpanAnchorsAt(sVisible, cursor, sAnchor, lits(i + 1), lCmp) Then
                sSpan = Mid(sVisible, cursor, Len(sAnchor))
                sLit = Mid(sVisible, cursor + Len(sAnchor), Len(lits(i + 1)))
                cursor = cursor + Len(sAnchor) + Len(lits(i + 1))
                bSpanFound = True
            End If
        End If

        If Not bSpanFound Then
            If Len(lits(i + 1)) = 0 Then
                ' Trailing empty literal: the span runs to end-of-string.
                sSpan = Mid(sVisible, cursor)
                sLit = ""
                cursor = Len(sVisible) + 1
            Else
                pos = InStr(cursor, sVisible, lits(i + 1), vbTextCompare)
                If pos = 0 Then Exit Function          ' literal lost -> conservative fallback
                sSpan = Mid(sVisible, cursor, pos - cursor)
                sLit = Mid(sVisible, pos, Len(lits(i + 1)))
                cursor = pos + Len(lits(i + 1))
            End If
        End If

        bSurvived = False
        If bHasEntry Then
            If Len(sEntry) > 0 Then
                ' A VALUE is compared BINARY: values are copied verbatim, so their casing is data.
                bSurvived = (sSpan = sEntry)
            Else
                ' Unset state: the literal token IS the last rendering, so finding it intact means the
                ' user left the token alone (a static-text edit must not release it). Case-insensitive,
                ' like the tokeniser: TokenLiteral re-materialises the canonical "Prop[" prefix, which is
                ' not necessarily the casing the user typed and the Template stored.
                bSurvived = (StrComp(sSpan, TokenLiteral(toks(i)), vbTextCompare) = 0)
            End If
        End If

        ' D7 - the span still CONTAINS the value intact, with the user's text beside it: may that addition be
        ' kept as static content while the token stays live? Each test is handed the addition concatenated
        ' with its context (the neighbouring literal), never the addition alone. Full rationale (the two
        ' forges a context-free test would miss): see "D7" in property-rendering-mechanics.md.
        sKeep = ""
        sKeepAfter = ""
        If Not bSurvived And bHasEntry Then
            If Len(sEntry) > 0 And Len(sSpan) > Len(sEntry) Then
                If Right(sSpan, Len(sEntry)) = sEntry Then
                    sKeep = Left(sSpan, Len(sSpan) - Len(sEntry))
                    If PrefixIsSafeAddition(lits(i) & sKeep, sEntry) Then
                        bSurvived = True
                    Else
                        sKeep = ""
                    End If
                ElseIf Left(sSpan, Len(sEntry)) = sEntry Then
                    sKeepAfter = Mid(sSpan, Len(sEntry) + 1)
                    If SuffixIsSafeAddition(sEntry, sKeepAfter & lits(i + 1)) Then
                        bSurvived = True
                    Else
                        sKeepAfter = ""
                    End If
                End If
            End If
        End If

        If bSurvived Then
            ' Both are empty on every path except an accepted addition, where they carry the user's own
            ' text back into the Template as static content - on the correct side of the token, which is
            ' what makes the re-authored Template render the next value in the place the user chose.
            sOut = sOut & sKeep & TokenLiteral(toks(i)) & sKeepAfter
            AppendValue toks(i), sEntry, NewNames, NewValues, nNew
        Else
            sOut = sOut & sSpan
        End If

        sOut = sOut & sLit
    Next i

    ' Anything the user appended past the final literal is kept as static text.
    If cursor <= Len(sVisible) Then sOut = sOut & Mid(sVisible, cursor)

    sNewTemplate = sOut
    AlignVisible = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.AlignVisible"
    AlignVisible = False
    nNew = 0
End Function

' True when sAnchor - what the LAST rendering left in this span - still sits exactly at the cursor AND is
' immediately followed by the closing literal. D6 relaxes the trailing-empty-literal case in one direction
' (safe appended text). Full rationale: see "SpanAnchorsAt" in property-rendering-mechanics.md.
Private Function SpanAnchorsAt(ByVal sVisible As String, ByVal cursor As Long, ByVal sAnchor As String, ByVal sNextLit As String, ByVal lAnchorCompare As Long) As Boolean
    SpanAnchorsAt = False
    If Len(sAnchor) = 0 Then Exit Function
    If cursor < 1 Then Exit Function
    If cursor + Len(sAnchor) - 1 > Len(sVisible) Then Exit Function
    If StrComp(Mid(sVisible, cursor, Len(sAnchor)), sAnchor, lAnchorCompare) <> 0 Then Exit Function

    If Len(sNextLit) = 0 Then
        If cursor + Len(sAnchor) = Len(sVisible) + 1 Then
            SpanAnchorsAt = True
        Else
            SpanAnchorsAt = SuffixIsSafeAddition(sAnchor, Mid(sVisible, cursor + Len(sAnchor)))
        End If
    Else
        SpanAnchorsAt = (StrComp(Mid(sVisible, cursor + Len(sAnchor), Len(sNextLit)), sNextLit, vbTextCompare) = 0)
    End If
End Function

' D6 - ASCII letter / ASCII digit, by CODE POINT (AscW, not Asc - locale-independent). An accented letter,
' a Unicode digit or a fullwidth digit is none of these, and is always the REFUSING side below.
Private Function IsAsciiLetter(ByVal sChar As String) As Boolean
    Dim n As Long
    IsAsciiLetter = False
    If Len(sChar) <> 1 Then Exit Function
    n = AscW(sChar)
    IsAsciiLetter = (n >= 65 And n <= 90) Or (n >= 97 And n <= 122)
End Function

Private Function IsAsciiDigit(ByVal sChar As String) As Boolean
    Dim n As Long
    IsAsciiDigit = False
    If Len(sChar) <> 1 Then Exit Function
    n = AscW(sChar)
    IsAsciiDigit = (n >= 48 And n <= 57)
End Function

' D6 - the whitelist test: True only for a character explicitly admitted (ASCII letter, or a
' SAFE_BOUNDARY_SYMBOLS member). Refusal is the DEFAULT - see property-rendering-mechanics.md.
Private Function IsSafeBoundaryChar(ByVal sChar As String) As Boolean
    On Error GoTo ErrorHandler

    IsSafeBoundaryChar = False
    If Len(sChar) <> 1 Then Exit Function

    If IsAsciiLetter(sChar) Then
        IsSafeBoundaryChar = True
    Else
        IsSafeBoundaryChar = (InStr(1, SAFE_BOUNDARY_SYMBOLS, sChar, vbBinaryCompare) > 0)
    End If
    Exit Function

ErrorHandler:
    IsSafeBoundaryChar = False
End Function

' D6 - may an addition AFTER the value be kept while the token keeps updating? Three load-bearing
' conditions (numeric anchor, whitelisted boundary char, exponent guard) that bound the CONSEQUENCE of a
' wrong call rather than guess intent. sSuffix is EVERYTHING to the right of the value, not just what the
' user typed. Full rationale: see "D6 - SuffixIsSafeAddition" in property-rendering-mechanics.md.
Private Function SuffixIsSafeAddition(ByVal sAnchor As String, ByVal sSuffix As String) As Boolean
    On Error GoTo ErrorHandler
    Dim s As String
    Dim k As Long

    SuffixIsSafeAddition = False
    If Len(sSuffix) = 0 Then Exit Function
    If Not StringsInEl.IsNumericText(sAnchor) Then Exit Function
    If Len(sAnchor) = 0 Then Exit Function          ' IsNumericText("") is True - never let it through

    s = LTrim(sSuffix)
    If Len(s) = 0 Then Exit Function                ' whitespace only: fuses with the next value
    If Not IsSafeBoundaryChar(Left(s, 1)) Then Exit Function

    If IsAsciiLetter(Left(s, 1)) Then
        k = 2
        If Mid(s, k, 1) = "+" Or Mid(s, k, 1) = "-" Then k = k + 1
        If IsAsciiDigit(Mid(s, k, 1)) Then Exit Function
    End If

    SuffixIsSafeAddition = True
    Exit Function

ErrorHandler:
    SuffixIsSafeAddition = False
End Function

' D6 mirror - may an addition BEFORE the value be kept? Same three conditions, opposite end. sPrefix is
' EVERYTHING to the left of the value, not just what the user typed. Full rationale: see "D6 mirror -
' PrefixIsSafeAddition" in property-rendering-mechanics.md.
Private Function PrefixIsSafeAddition(ByVal sPrefix As String, ByVal sAnchor As String) As Boolean
    On Error GoTo ErrorHandler
    Dim s As String

    PrefixIsSafeAddition = False
    If Len(sPrefix) = 0 Then Exit Function
    If Not StringsInEl.IsNumericText(sAnchor) Then Exit Function
    If Len(sAnchor) = 0 Then Exit Function

    s = RTrim(sPrefix)
    If Len(s) = 0 Then Exit Function
    If Not IsSafeBoundaryChar(Right(s, 1)) Then Exit Function

    If IsAsciiLetter(Right(s, 1)) And Len(s) >= 2 Then
        If IsAsciiDigit(Mid(s, Len(s) - 1, 1)) Then Exit Function
    End If

    PrefixIsSafeAddition = True
    Exit Function

ErrorHandler:
    PrefixIsSafeAddition = False
End Function

' D8 - the RAW frontier of a literal (first/last two characters, BEFORE any trimming): did the text
' touching the value change? Full rationale: see "D8" in property-rendering-mechanics.md.
Private Function RawHead(ByVal s As String) As String
    RawHead = Left(s, 2)
End Function

Private Function RawTail(ByVal s As String) As String
    RawTail = Right(s, 2)
End Function

' D8 - could the END of this literal be the beginning of a number? Answered on the trailing run of
' number-capable characters taken WHOLE, bounded by the data never a constant: a number starts with a
' digit. Closes the non-local base-literal forge D6's fixed-distance guards cannot see. Full rationale:
' see "D8" in property-rendering-mechanics.md.
Private Function NumericTailIsPossible(ByVal sLit As String) As Boolean
    On Error GoTo ErrorHandler
    Dim i As Long

    NumericTailIsPossible = False
    i = Len(sLit)
    Do While i > 0
        If InStr(1, NUMBER_CAPABLE_CHARS, Mid(sLit, i, 1), vbBinaryCompare) = 0 Then Exit Do
        i = i - 1
    Loop

    If i >= Len(sLit) Then Exit Function           ' no number-capable run at all
    NumericTailIsPossible = IsAsciiDigit(Mid(sLit, i + 1, 1))
    Exit Function

ErrorHandler:
    NumericTailIsPossible = False
End Function

' D8 - is the text now sitting to the LEFT of a value acceptable? Three ordered questions (new base
' literal possible? did the RAW frontier change at all? then D6's guard). Full rationale: see "D8" in
' property-rendering-mechanics.md.
Private Function LeftContextIsSafe(ByVal sAuthored As String, ByVal sCurrent As String, ByVal sValue As String) As Boolean
    LeftContextIsSafe = True

    ' EMPTY current literal provably cannot weld (nothing to weld with) - handled HERE, never in the shared
    ' D6 predicates, where empty means the opposite ("no addition was made").
    If Len(sCurrent) = 0 Then Exit Function

    LeftContextIsSafe = False
    If NumericTailIsPossible(sCurrent) And Not NumericTailIsPossible(sAuthored) Then Exit Function

    LeftContextIsSafe = True
    If RawTail(sAuthored) = RawTail(sCurrent) Then Exit Function

    LeftContextIsSafe = PrefixIsSafeAddition(sCurrent, sValue)
End Function

' D8 mirror, for the text now sitting to the RIGHT of a value - no base-literal guard needed (that family
' can only form to the LEFT). See property-rendering-mechanics.md.
Private Function RightContextIsSafe(ByVal sAuthored As String, ByVal sCurrent As String, ByVal sValue As String) As Boolean
    RightContextIsSafe = True

    ' Emptied literal (nothing to weld) handled here, same as the left half. An emptied gap BETWEEN two
    ' values is caught by ReauthoredTemplateIsSound's well-formedness check instead.
    If Len(sCurrent) = 0 Then Exit Function
    If RawHead(sAuthored) = RawHead(sCurrent) Then Exit Function

    RightContextIsSafe = SuffixIsSafeAddition(sValue, sCurrent)
End Function

' D8 - the ONE check that makes a re-authored Template safe to store: well-formed, carries exactly the
' expected tokens, and expanding it with the SAME LastValues reproduces the visible text byte-for-byte
' (the engine's fixed point, verified not inherited). Any failure falls back conservatively. Full
' rationale (the round-8/10 wedge case): see property-rendering-mechanics.md.
Private Function ReauthoredTemplateIsSound(ByVal sNewTemplate As String, ByVal nExpectedTok As Long, ByRef names() As String, ByRef values() As String, ByVal n As Long, ByVal sVisible As String, ByVal bValidateNames As Boolean) As Boolean
    On Error GoTo ErrorHandler

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim bOk As Boolean

    ReauthoredTemplateIsSound = False

    If Not TemplateIsWellFormed(sNewTemplate, bValidateNames) Then Exit Function
    If Not ParseTemplate(sNewTemplate, lits, toks, nTok, bValidateNames) Then Exit Function
    If nTok <> nExpectedTok Then Exit Function

    bOk = False
    If ExpandTemplate(sNewTemplate, names, values, n, bValidateNames, bOk) <> sVisible Then Exit Function
    If Not bOk Then Exit Function

    ReauthoredTemplateIsSound = True
    Exit Function

ErrorHandler:
    ReauthoredTemplateIsSound = False
End Function

' D8 - alignment by VALUE RECOGNITION, the fallback for everything the literal walk cannot follow. Runs
' ONLY after AlignVisible has declined. Public test seam (AC12). See "D8 - AlignByValues" in
' property-rendering-mechanics.md.
Public Function AlignByValues(ByVal sVisible As String, ByVal sTemplate As String, ByRef ValNames() As String, ByRef ValValues() As String, ByVal nVals As Long, ByRef sNewTemplate As String, ByRef NewNames() As String, ByRef NewValues() As String, ByRef nNew As Long, Optional ByVal bValidateNames As Boolean = True) As Boolean
    On Error GoTo ErrorHandler

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim i As Long
    Dim pos As Long
    Dim cursor As Long
    Dim sEntry As String
    Dim sPrevEntry As String
    Dim sGap As String
    Dim sOut As String

    AlignByValues = False
    sNewTemplate = ""
    nNew = 0

    If Not ParseTemplate(sTemplate, lits, toks, nTok, bValidateNames) Then Exit Function
    If nTok = 0 Then Exit Function

    cursor = 1
    sOut = ""
    sPrevEntry = ""

    For i = 0 To nTok - 1
        If Not HasValueEntry(toks(i), ValNames, nVals) Then Exit Function
        sEntry = LookupValue(toks(i), ValNames, ValValues, nVals)
        If Len(sEntry) = 0 Then Exit Function

        ' A VALUE is matched BINARY, exactly as AlignVisible does: values are copied verbatim, so casing
        ' is data and a case-insensitive hit would align the wrong text.
        pos = InStr(cursor, sVisible, sEntry, vbBinaryCompare)
        If pos = 0 Then Exit Function
        If InStr(pos + 1, sVisible, sEntry, vbBinaryCompare) > 0 Then Exit Function

        sGap = Mid(sVisible, cursor, pos - cursor)

        ' The gap touches TWO values - the previous one on its left edge, this one on its right - and both
        ' edges must be judged, against the authored literal that used to sit there.
        If i > 0 Then
            If Not RightContextIsSafe(lits(i), sGap, sPrevEntry) Then Exit Function
        End If
        If Not LeftContextIsSafe(lits(i), sGap, sEntry) Then Exit Function

        sOut = sOut & sGap & TokenLiteral(toks(i))
        AppendValue toks(i), sEntry, NewNames, NewValues, nNew
        cursor = pos + Len(sEntry)
        sPrevEntry = sEntry
    Next i

    sGap = Mid(sVisible, cursor)
    If Not RightContextIsSafe(lits(nTok), sGap, sPrevEntry) Then Exit Function
    sOut = sOut & sGap

    If Not ReauthoredTemplateIsSound(sOut, nTok, NewNames, NewValues, nNew, sVisible, bValidateNames) Then Exit Function

    sNewTemplate = sOut
    AlignByValues = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.AlignByValues"
    AlignByValues = False
    nNew = 0
End Function

'######################################################################################################################
'                                          STATE MACHINE
'######################################################################################################################

' One drain target: skip it when this drain already ran it (done-list dedup), re-fetch it, then run the
' state machine. Full rationale (the "two bound texts in one group" failure this prevents): see "Repaint
' hop (bounded)" in property-rendering-mechanics.md.
Private Sub RenderDrainTarget(ByVal oEl As element, ByRef doneIds() As String, ByRef nDone As Long)
    On Error GoTo ErrorHandler

    Dim sKey As String
    Dim oFresh As element

    If oEl Is Nothing Then Exit Sub

    ' An element whose id cannot be read simply gets no dedup - it may be rendered twice, but it is never
    ' wrongly skipped.
    sKey = ElementIdKey(oEl)
    If Len(sKey) > 0 Then
        If IsIdDone(sKey, doneIds, nDone) Then Exit Sub
        ReDim Preserve doneIds(0 To nDone)
        doneIds(nDone) = sKey
        nDone = nDone + 1
    End If

    Set oFresh = FreshHandle(oEl)
    If oFresh Is Nothing Then Exit Sub
    RenderOneElement oFresh
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.RenderDrainTarget"
End Sub

' Re-fetch an element from the model so the state machine reads what is REALLY in the file. THE rule of
' this module: never read text through a handle you did not just fetch. Full rationale: see "The
' stale-handle invariant (FreshHandle)" in property-rendering-mechanics.md.
Private Function FreshHandle(ByVal oEl As element) As element
    On Error Resume Next

    Dim oNew As element

    Set FreshHandle = oEl
    If oEl Is Nothing Then Exit Function
    Set oNew = ActiveModelReference.GetElementById(oEl.id)
    If Not oNew Is Nothing Then Set FreshHandle = oNew
End Function

' Stable per-element key for the drain's done-list. DLongToString is the sanctioned conversion - an
' Element.ID is a DLong and must never be used as a plain value - and it is the same key ElementInProcesse
' uses to identify an element across a batch. "" when the id cannot be read.
Private Function ElementIdKey(ByVal oEl As element) As String
    On Error Resume Next
    ElementIdKey = ""
    ElementIdKey = DLongToString(oEl.id)
End Function

Private Function IsIdDone(ByVal sKey As String, ByRef doneIds() As String, ByVal nDone As Long) As Boolean
    Dim i As Long
    IsIdDone = False
    For i = 0 To nDone - 1
        If doneIds(i) = sKey Then
            IsIdDone = True
            Exit Function
        End If
    Next i
End Function

' Run the full state machine on one element, whatever its binding state. The repaint hop's only entry
' point, and the reason no blind render exists anywhere in this module.
Private Sub RenderOneElement(ByVal oEl As element)
    On Error GoTo ErrorHandler

    If oEl Is Nothing Then Exit Sub
    If mbWriteDisabled Then Exit Sub
    If Not IsRenderableType(oEl) Then Exit Sub
    If oEl.IsLocked Then Exit Sub
    If Not IsSysLibraryPresent() Then Exit Sub

    Dim ids() As Long
    Dim texts() As String
    Dim nBearers As Long
    nBearers = StringsInEl.EnumerateTextSubIds(oEl, ids, texts)
    If nBearers <= 0 Then Exit Sub

    If CustomPropertyHandler.IsItemAttachedToElement(oEl, ARES_ITEM_RENDER, ARES_NAME_LIBRARY_SYS) Then
        RenderBoundElement oEl, texts, nBearers
    Else
        TryFirstAuthor oEl, texts, nBearers, False
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.RenderOneElement"
End Sub

' Branches 1-3, per stored entry. texts() holds the whole text of every bearer, indexed by SubId. Write
' ORDER is load-bearing: every TEXT write happens first, the ONE metadata write last. Full rationale: see
' "RenderBoundElement / RenderEntryOnElement" in property-rendering-mechanics.md.
Private Sub RenderBoundElement(ByRef oEl As element, ByRef texts() As String, ByVal nBearers As Long)
    On Error GoTo ErrorHandler

    Dim ents() As RenderEntry
    Dim nEnts As Long
    Dim i As Long
    Dim outcome As Long
    Dim bDirty As Boolean
    Dim nLive As Long
    Dim wrIds() As Long
    Dim wrPrev() As String
    Dim nWr As Long
    Dim wSubId As Long
    Dim wPrev As String
    Dim subIds() As Long
    Dim nAdded As Long
    Dim bRepaint As Boolean

    ' ReadRenderMetadata surfaces the right status itself (a schema mismatch has a dedicated one, and the
    ' generic "unreadable" must not overwrite it).
    If Not ReadRenderMetadata(oEl, ents, nEnts) Then Exit Sub

    ' Every entry is located up front, in two passes over the whole list, so that one sub-text is driven
    ' by ONE entry only and the result does not depend on the order the entries happen to sit in.
    If nEnts > 0 Then ResolveAllSubIds ents, nEnts, texts, nBearers, subIds

    ' THEN top up: a sub-text that no entry claims has never been authored, and being on an element that
    ' is already bound is not a reason to ignore it. Binding is per SUB-TEXT; the item on the cell header
    ' is just where the whole list lives. This also covers the legal "attached but Entries empty" state.
    nAdded = AuthorUnclaimedBearers(oEl, texts, nBearers, ents, nEnts, subIds)
    If nEnts = 0 Then Exit Sub
    If nAdded > 0 Then ResolveAllSubIds ents, nEnts, texts, nBearers, subIds

    ' A newly authored entry must be PERSISTED even when its own render is a no-op (an unset value renders
    ' the literal token, which is what the bearer already shows - branch 1). Without this the new entry
    ' would be re-authored, and thrown away, on every single pass.
    bDirty = (nAdded > 0)
    nWr = 0
    For i = 0 To nEnts - 1
        wSubId = -1
        wPrev = ""
        outcome = RenderEntryOnElement(oEl, ents, i, texts, nBearers, subIds(i), wSubId, wPrev, bRepaint)
        If wSubId >= 0 Then
            ReDim Preserve wrIds(0 To nWr)
            ReDim Preserve wrPrev(0 To nWr)
            wrIds(nWr) = wSubId
            wrPrev(nWr) = wPrev
            nWr = nWr + 1
        End If
        If outcome = ENTRY_DROP Then
            ents(i).Dropped = True
            bDirty = True
        ElseIf outcome = ENTRY_UPDATED Then
            bDirty = True
        End If
    Next i

    If Not bDirty Then Exit Sub

    ' The whole binding was released: take the metadata off rather than leave an empty shell behind.
    nLive = CountLiveEntries(ents, nEnts)
    If nLive = 0 Then
        PropertyTagging.DetachRenderMetadata oEl
        Exit Sub
    End If

    If Not WriteRenderMetadata(oEl, ents, nEnts) Then
        For i = 0 To nWr - 1
            StringsInEl.SetTextAtSubId oEl, wrIds(i), wrPrev(i)
        Next i
        ReportMetadataUnreadable
        DisableAfterWriteFailure
    ElseIf bRepaint Then
        ' Branch 3 transitioned an entry WITHOUT writing text. Noted only HERE, after the metadata write
        ' succeeded - see "RenderBoundElement / RenderEntryOnElement" in property-rendering-mechanics.md.
        NoteDirtyGroup oEl
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.RenderBoundElement"
End Sub

' The three branches for ONE entry. Returns ENTRY_UNCHANGED / ENTRY_UPDATED / ENTRY_DROP. SubId is the
' ordinal ResolveAllSubIds settled (-1 = refuse; a drift relocation may differ from the stored one). oEl is
' ByRef down to StringsInEl.SetTextAtSubId (see WriteRenderedText). Full rationale for each branch: see
' "RenderBoundElement / RenderEntryOnElement" in property-rendering-mechanics.md.
Private Function RenderEntryOnElement(ByRef oEl As element, ByRef ents() As RenderEntry, ByVal idx As Long, ByRef texts() As String, ByVal nBearers As Long, ByVal SubId As Long, ByRef wSubId As Long, ByRef wPrev As String, ByRef bRepaint As Boolean) As Long
    On Error GoTo ErrorHandler

    Dim sVisible As String
    Dim sExpCur As String
    Dim sExpLast As String
    Dim bOkCur As Boolean
    Dim bOkLast As Boolean
    Dim curNames() As String
    Dim curValues() As String
    Dim nCur As Long
    Dim newNames() As String
    Dim newValues() As String
    Dim nNew As Long
    Dim sNewTemplate As String
    Dim newLits() As String
    Dim newToks() As String
    Dim nNewTok As Long
    Dim k As Long
    Dim bRelocated As Boolean

    RenderEntryOnElement = ENTRY_UNCHANGED
    wSubId = -1
    wPrev = ""

    ' The sub-text could not be identified (its ordinal drifted and its text is nowhere to be found, or
    ' another entry owns it): refuse rather than write blind into a sub-text that is not ours.
    If SubId < 0 Then
        ReportDrift
        Exit Function
    End If
    If SubId > nBearers - 1 Then
        ReportDrift
        Exit Function
    End If
    ' Relocation held in a LOCAL and committed only once the entry proves usable - see
    ' property-rendering-mechanics.md.
    bRelocated = (SubId <> ents(idx).SubId)

    sVisible = texts(SubId)
    ' An empty or faulted read is NEVER "the user wiped everything": skip, no transition.
    If Len(sVisible) = 0 Then
        ReportMetadataUnreadable
        Exit Function
    End If

    ' Validation on read: the Template's token set must match the LastValues key set exactly. Metadata
    ' vandalised through the native Properties pane is never rendered as if it had been intended.
    If Not EntryIsConsistent(ents(idx)) Then
        ReportMetadataInvalid
        Exit Function
    End If

    If Not ReadCurrentValues(oEl, ents(idx), curNames, curValues, nCur) Then Exit Function

    ' ExpandTemplate fails OPEN and reaches COM - skip the entry on a fault, exactly as an empty read does.
    ' See "ExpandTemplate" in property-rendering-mechanics.md.
    sExpCur = ExpandTemplate(ents(idx).Template, curNames, curValues, nCur, True, bOkCur)
    If Not bOkCur Then Exit Function
    sExpLast = ExpandTemplate(ents(idx).Template, ents(idx).ValNames, ents(idx).ValValues, ents(idx).nVals, True, bOkLast)
    If Not bOkLast Then Exit Function

    ' Relocation is persisted only once the entry is proven readable/consistent/legal (edge #17). An
    ' undrifted entry still returns ENTRY_UNCHANGED, so branch 1 stays the strict no-op AC2 requires.
    If bRelocated Then
        ents(idx).SubId = SubId
        RenderEntryOnElement = ENTRY_UPDATED
    End If

    ' --- BRANCH 1: up to date. STRICT no-op - no text write, no Rewrite, no metadata write. The loop
    ' terminator, and the reason a re-queued unchanged element costs nothing.
    If sVisible = sExpCur Then Exit Function

    ' --- BRANCH 2: the visible still matches the LAST rendering, so the VALUES moved. Re-render. Second
    ' test also catches an entry authored but never rendered (branch 3 trap otherwise) - full rationale:
    ' see "RenderBoundElement / RenderEntryOnElement" in property-rendering-mechanics.md.
    If sVisible = sExpLast Or StrComp(sVisible, ents(idx).Template, vbTextCompare) = 0 Then
        If WriteRenderedText(oEl, SubId, sExpCur) Then
            wSubId = SubId
            wPrev = sVisible
            SetEntryValues ents, idx, curNames, curValues, nCur
            RenderEntryOnElement = ENTRY_UPDATED
        End If
        Exit Function
    End If

    ' --- BRANCH 3: a NON-EMPTY read differing from BOTH expansions - positive proof of a user edit.
    ' Per-token release, no text written. Full rationale: see "RenderBoundElement / RenderEntryOnElement"
    ' in property-rendering-mechanics.md.
    If AlignVisible(sVisible, ents(idx).Template, ents(idx).ValNames, ents(idx).ValValues, ents(idx).nVals, sNewTemplate, newNames, newValues, nNew) Then
        ' The re-authored Template is re-validated exactly like a first author.
        If TemplateIsWellFormed(sNewTemplate, True) Then
            If ParseTemplate(sNewTemplate, newLits, newToks, nNewTok, True) Then
                If nNewTok > 0 Then
                    For k = 0 To nNewTok - 1
                        If Not HasValueEntry(newToks(k), newNames, nNew) Then
                            AppendValue newToks(k), "", newNames, newValues, nNew
                        End If
                    Next k
                    ents(idx).Template = sNewTemplate
                    SetEntryValues ents, idx, newNames, newValues, nNew
                    bRepaint = True
                    RenderEntryOnElement = ENTRY_UPDATED
                Else
                    ' Alignment SUCCEEDED and concluded nothing survived - a deliberate release, not a
                    ' failure to understand the text; gets its own status rather than ReportAmbiguous below.
                    ReportBindingReleased
                    RenderEntryOnElement = ENTRY_DROP
                End If
                Exit Function
            End If
        End If
    End If

    ' --- BRANCH 3b (D8): the literal walk could not follow the text (a position moved anywhere). Try to
    ' recognise the values themselves instead. Runs ONLY after AlignVisible has declined.
    If AlignByValues(sVisible, ents(idx).Template, ents(idx).ValNames, ents(idx).ValValues, ents(idx).nVals, sNewTemplate, newNames, newValues, nNew) Then
        ents(idx).Template = sNewTemplate
        SetEntryValues ents, idx, newNames, newValues, nNew
        bRepaint = True
        RenderEntryOnElement = ENTRY_UPDATED
        Exit Function
    End If

    ' Ambiguous alignment (or an invalid re-authored Template): keep only the literal tokens the visible
    ' still carries, drop everything else, and say so.
    ReportAmbiguous
    RenderEntryOnElement = ApplyConservativeFallback(ents, idx, sVisible)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.RenderEntryOnElement"
    RenderEntryOnElement = ENTRY_UNCHANGED
End Function

' Conservative outcome: the visible text BECOMES the Template. Converges next pass (branch 2 or 1); an
' ill-formed or token-free visible RELEASES the entry outright instead. Full rationale: see
' "ApplyConservativeFallback" in property-rendering-mechanics.md.
Private Function ApplyConservativeFallback(ByRef ents() As RenderEntry, ByVal idx As Long, ByVal sVisible As String) As Long
    On Error GoTo ErrorHandler

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim i As Long
    Dim names() As String
    Dim values() As String
    Dim n As Long

    ApplyConservativeFallback = ENTRY_DROP

    If Not TemplateIsWellFormed(sVisible, True) Then Exit Function
    If Not ParseTemplate(sVisible, lits, toks, nTok, True) Then Exit Function
    If nTok = 0 Then Exit Function

    n = 0
    For i = 0 To nTok - 1
        AppendValue toks(i), "", names, values, n
    Next i

    ents(idx).Template = sVisible
    SetEntryValues ents, idx, names, values, n
    ApplyConservativeFallback = ENTRY_UPDATED
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.ApplyConservativeFallback"
    ApplyConservativeFallback = ENTRY_DROP
End Function

' Author ONE bearer, if it qualifies, appending an entry to the list. THE unit of binding is the SUB-TEXT,
' not the element. Both the first-author scan and the top-up scan run this one function, so the rules
' cannot drift. Full rationale (hybrid policy §6.2, source-preservation via bKeepSource/nFree): see
' "Authoring (TryAuthorBearer / AuthorUnclaimedBearers / TryFirstAuthor)" in property-rendering-mechanics.md.
Private Function TryAuthorBearer(ByRef oEl As element, ByVal sText As String, ByVal SubId As Long, ByRef ents() As RenderEntry, ByRef nEnts As Long, ByVal bKeepSource As Boolean, ByRef nFree As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim names() As String
    Dim values() As String
    Dim n As Long
    Dim j As Long
    Dim bAllAttached As Boolean

    TryAuthorBearer = False

    If Not ParseTemplate(sText, lits, toks, nTok, True) Then Exit Function
    If nTok = 0 Then Exit Function

    ' SOURCE PRESERVATION, after the token test on purpose - see property-rendering-mechanics.md.
    If bKeepSource Then
        If nFree <= 1 Then
            ReportCycleWarning
            Exit Function
        End If
    End If

    If Not TemplateIsWellFormed(sText, True) Then
        ' Duplicate or adjacent tokens: refused at author time, no metadata written.
        ReportStructuralRefusal sText
        Exit Function
    End If
    If Not CanBearTokens(oEl, sText) Then
        ReportTextNodeRefused
        Exit Function
    End If

    bAllAttached = True
    n = 0
    For j = 0 To nTok - 1
        If Not CustomPropertyHandler.IsItemAttachedToElement(oEl, toks(j)) Then
            bAllAttached = False
        Else
            AppendValue toks(j), "", names, values, n
            WarnGovernedValue oEl, toks(j)
        End If
    Next j

    If Not bAllAttached Then
        ReportNotBound
        Exit Function
    End If

    ReDim Preserve ents(0 To nEnts)
    ents(nEnts).SubId = SubId
    ents(nEnts).Template = sText
    ents(nEnts).Dropped = False
    SetEntryValues ents, nEnts, names, values, n
    nEnts = nEnts + 1
    nFree = nFree - 1
    TryAuthorBearer = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.TryAuthorBearer"
    TryAuthorBearer = False
End Function

' Does this element FEED a Cell* calc source? If so it must keep at least ONE sub-text out of the exclusion
' set (TryAuthorBearer enforces it) - binding the last one would let the containment destroy the data it
' protects. Full rationale: see "Authoring" in property-rendering-mechanics.md.
Private Function FeedsCellSource(ByRef oEl As element) As Boolean
    On Error GoTo ErrorHandler

    FeedsCellSource = False
    If oEl Is Nothing Then Exit Function
    If Not oEl.IsCellElement Then Exit Function
    FeedsCellSource = PropertyCalculation.IsTriggerCell(oEl)
    Exit Function

ErrorHandler:
    ' Fail-closed: if we cannot tell, assume the cell feeds a source and protect it.
    FeedsCellSource = True
End Function

' TOP-UP AUTHORING on a bearer that is ALREADY bound: every sub-text no stored entry claims is offered to
' the same authoring rule. Returns how many entries were added. Without this a second sub-text whose
' property arrives later stays inert forever. Full rationale: see "Authoring" in
' property-rendering-mechanics.md.
Private Function AuthorUnclaimedBearers(ByRef oEl As element, ByRef texts() As String, ByVal nBearers As Long, ByRef ents() As RenderEntry, ByRef nEnts As Long, ByRef subIds() As Long) As Long
    On Error GoTo ErrorHandler

    Dim claimed() As Boolean
    Dim i As Long
    Dim nBefore As Long
    Dim nFree As Long
    Dim bKeepSource As Boolean

    AuthorUnclaimedBearers = 0
    If nBearers <= 0 Then Exit Function
    If Not IsAcceptableBearerElement(oEl) Then Exit Function

    ' An UNRESOLVED live entry (subIds = -1) bails the whole scan this pass, rather than risk authoring a
    ' duplicate for the text it drives. See "Authoring" in property-rendering-mechanics.md.
    ReDim claimed(0 To nBearers - 1)
    For i = 0 To nEnts - 1
        If Not ents(i).Dropped Then
            If subIds(i) < 0 Then Exit Function
            If subIds(i) <= nBearers - 1 Then claimed(subIds(i)) = True
        End If
    Next i

    nFree = 0
    For i = 0 To nBearers - 1
        If Not claimed(i) Then nFree = nFree + 1
    Next i
    If nFree = 0 Then Exit Function

    ' Evaluated once, and only once there is something to author, so a fully-bound cell pays nothing.
    ' TryAuthorBearer owns the rule itself and maintains nFree.
    bKeepSource = FeedsCellSource(oEl)

    nBefore = nEnts
    For i = 0 To nBearers - 1
        If Not claimed(i) Then TryAuthorBearer oEl, texts(i), i, ents, nEnts, bKeepSource, nFree
    Next i

    AuthorUnclaimedBearers = nEnts - nBefore
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.AuthorUnclaimedBearers"
    AuthorUnclaimedBearers = 0
End Function

' Branch 4 - the first author, hybrid policy. bManual = the BindPropertyRender key-in (reports refusals);
' the automatic path uses identical rules. A bearer binds only when every token names a property ALREADY
' ATTACHED to the element - attachment is the intent signal, so nothing is ever bound by accident.
Private Function TryFirstAuthor(ByRef oEl As element, ByRef texts() As String, ByVal nBearers As Long, ByVal bManual As Boolean) As Boolean
    On Error GoTo ErrorHandler

    Dim ents() As RenderEntry
    Dim nEnts As Long
    Dim i As Long
    Dim nFree As Long
    Dim bKeepSource As Boolean
    Dim sExp As String
    Dim bOkExp As Boolean
    Dim curNames() As String
    Dim curValues() As String
    Dim nCur As Long
    Dim wrIds() As Long
    Dim wrPrev() As String
    Dim nWr As Long

    TryFirstAuthor = False
    nEnts = 0

    ' Element-level bearer guard, evaluated once: never author on anything but a TOP-LEVEL model element.
    If Not IsAcceptableBearerElement(oEl) Then Exit Function

    ' One decision per BEARER, via the shared authoring rule (includes SOURCE PRESERVATION - see
    ' "Authoring" in property-rendering-mechanics.md). Nothing is bound yet, so every bearer starts free.
    bKeepSource = FeedsCellSource(oEl)
    nFree = nBearers
    For i = 0 To nBearers - 1
        TryAuthorBearer oEl, texts(i), i, ents, nEnts, bKeepSource, nFree
    Next i

    If nEnts = 0 Then Exit Function

    ' Render every new entry BEFORE any metadata is touched (text first, then metadata).
    nWr = 0
    For i = 0 To nEnts - 1
        If ReadCurrentValues(oEl, ents(i), curNames, curValues, nCur) Then
            ' A faulted expansion returns the Template unchanged; binding on it would store LastValues
            ' that do not match what is actually visible, and the next pass would read a user edit that
            ' never happened. Drop the entry instead.
            sExp = ExpandTemplate(ents(i).Template, curNames, curValues, nCur, True, bOkExp)
            If Not bOkExp Then
                ents(i).Dropped = True
            ElseIf WriteRenderedText(oEl, ents(i).SubId, sExp) Then
                ReDim Preserve wrIds(0 To nWr)
                ReDim Preserve wrPrev(0 To nWr)
                wrIds(nWr) = ents(i).SubId
                wrPrev(nWr) = texts(ents(i).SubId)
                nWr = nWr + 1
                SetEntryValues ents, i, curNames, curValues, nCur
            Else
                ents(i).Dropped = True
            End If
        Else
            ents(i).Dropped = True
        End If
    Next i

    If CountLiveEntries(ents, nEnts) = 0 Then Exit Function

    If Not PropertyTagging.AttachRenderMetadata(oEl) Then
        RestoreWrittenTexts oEl, wrIds, wrPrev, nWr
        ReportLibraryMissing
        Exit Function
    End If

    If Not WriteRenderMetadata(oEl, ents, nEnts) Then
        ' The binding never landed: restore the text and undo our own attach, so no half-bound element
        ' is left behind and the next pass sees exactly the state it saw this time.
        RestoreWrittenTexts oEl, wrIds, wrPrev, nWr
        PropertyTagging.DetachRenderMetadata oEl
        ReportMetadataUnreadable
        DisableAfterWriteFailure
        Exit Function
    End If

    If bManual Then LangManager.ShowStatusT "RenderBindDone"
    TryFirstAuthor = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.TryFirstAuthor"
    TryFirstAuthor = False
End Function

' Put back the texts written earlier in this pass, after the metadata write failed. Best effort by
' definition - a failure here leaves the visible ahead of the stored state, which the state machine then
' reads as a user edit rather than as a corrupt binding.
' oEl is ByRef so SetTextAtSubId's post-Rewrite handle refresh reaches the caller (see WriteRenderedText).
Private Sub RestoreWrittenTexts(ByRef oEl As element, ByRef ids() As Long, ByRef prev() As String, ByVal n As Long)
    On Error Resume Next
    Dim i As Long
    For i = 0 To n - 1
        StringsInEl.SetTextAtSubId oEl, ids(i), prev(i)
    Next i
End Sub

' The ONE place visible text is written. Refuses a reserved serialisation delimiter (and a stray CR) before
' it reaches the file. Returns True only when the sub-text now reads sNew. oEl is ByRef down the whole call
' chain on purpose - see "WriteRenderedText and the ByRef element chain" in property-rendering-mechanics.md.
Private Function WriteRenderedText(ByRef oEl As element, ByVal SubId As Long, ByVal sNew As String) As Boolean
    On Error GoTo ErrorHandler

    WriteRenderedText = False

    ' vbLf is legal here (a multi-line TextNode's rendering carries it); only the delimiters and a stray
    ' CR are refused. VALUES are held to the stricter rule in ReadCurrentValues.
    If ContainsSerialisationDelimiter(sNew) Then
        ReportValueIllegal
        Exit Function
    End If
    If InStr(1, sNew, vbCr) > 0 Then
        ReportValueIllegal
        Exit Function
    End If

    WriteRenderedText = StringsInEl.SetTextAtSubId(oEl, SubId, sNew)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.WriteRenderedText"
    WriteRenderedText = False
End Function

' Read the CURRENT value of every token of an entry, off the bearing element's OWN attached properties.
' Values are copied VERBATIM (never CStr/Format/rounding). Full rationale: see "ReadCurrentValues" in
' property-rendering-mechanics.md.
Private Function ReadCurrentValues(ByVal oEl As element, ByRef ent As RenderEntry, ByRef names() As String, ByRef values() As String, ByRef n As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim i As Long
    Dim vVal As Variant
    Dim sVal As String

    ReadCurrentValues = False
    n = 0
    If Not ParseTemplate(ent.Template, lits, toks, nTok, True) Then Exit Function

    For i = 0 To nTok - 1
        sVal = ""
        vVal = CustomPropertyHandler.GetPropertyValueFromElement(oEl, toks(i), toks(i))
        ' Nested tests, never an And chain: VBA evaluates both operands, and VarType/CStr on an array
        ' would raise.
        If IsNull(vVal) Then
            sVal = ""                                   ' unset -> the literal token, per Expand
        ElseIf IsArray(vVal) Then
            ReportValueUnsupported
            sVal = ""
        ElseIf VarType(vVal) = vbString Then
            sVal = vVal                                 ' VERBATIM - no CStr, no Format, no rounding
        Else
            ReportValueUnsupported
            sVal = ""
        End If

        ' A VALUE is held to the strict rule: no delimiter, no CR, no LF. Refusing here is what makes it
        ' impossible for a value to render the metadata unparseable.
        If Len(sVal) > 0 Then
            If ValueHasIllegalChar(sVal) Then
                ReportValueIllegal
                Exit Function
            End If
        End If

        AppendValue toks(i), sVal, names, values, n
    Next i

    ReadCurrentValues = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.ReadCurrentValues"
    ReadCurrentValues = False
End Function

' Locate the sub-text EVERY entry drives, in TWO PASSES over the whole entry list (Pass A: certain matches
' only; Pass B: relocation scan). Independent of entry ORDER, and comparisons are BINARY (identification,
' not interpretation). Full rationale: see "SubId resolution - the two-pass algorithm" in
' property-rendering-mechanics.md.
Private Sub ResolveAllSubIds(ByRef ents() As RenderEntry, ByVal nEnts As Long, ByRef texts() As String, ByVal nBearers As Long, ByRef subIds() As Long)
    On Error GoTo ErrorHandler

    Dim claimed() As Boolean
    Dim i As Long

    ReDim subIds(0 To nEnts - 1)
    ReDim claimed(0 To nBearers - 1)

    For i = 0 To nEnts - 1
        subIds(i) = ResolveSubIdExact(ents, i, texts, nBearers, claimed)
    Next i

    For i = 0 To nEnts - 1
        If subIds(i) < 0 Then subIds(i) = ResolveSubIdRelocate(ents, i, texts, nBearers, claimed)
    Next i
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.ResolveAllSubIds"
    ' Refuse every entry rather than hand out a mapping that may be half-built.
    For i = 0 To nEnts - 1
        subIds(i) = -1
    Next i
End Sub

' Pass A: the CERTAIN case only. The text at the stored ordinal still is the computed last rendering, or
' the Template (an entry authored but not yet rendered). Nothing drifted, so the ordinal is claimed and
' no scan can take it away afterwards. -1 simply means "not certain here" - pass B decides.
Private Function ResolveSubIdExact(ByRef ents() As RenderEntry, ByVal idx As Long, ByRef texts() As String, ByVal nBearers As Long, ByRef claimed() As Boolean) As Long
    On Error GoTo ErrorHandler

    Dim SubId As Long

    ResolveSubIdExact = -1
    SubId = ents(idx).SubId
    If SubId < 0 Then Exit Function
    If SubId > nBearers - 1 Then Exit Function
    If IsSubIdClaimed(claimed, SubId) Then Exit Function

    If texts(SubId) = ExpandTemplate(ents(idx).Template, ents(idx).ValNames, ents(idx).ValValues, ents(idx).nVals) Then
        ResolveSubIdExact = ClaimSubId(claimed, SubId)
        Exit Function
    End If
    If texts(SubId) = ents(idx).Template Then ResolveSubIdExact = ClaimSubId(claimed, SubId)
    Exit Function

ErrorHandler:
    ResolveSubIdExact = -1
End Function

' Pass B, for entries pass A could not settle: relocate by matching text, else fall back to the stored
' ordinal, else refuse. Full rationale (edge #17): see "SubId resolution" in property-rendering-mechanics.md.
Private Function ResolveSubIdRelocate(ByRef ents() As RenderEntry, ByVal idx As Long, ByRef texts() As String, ByVal nBearers As Long, ByRef claimed() As Boolean) As Long
    On Error GoTo ErrorHandler

    Dim sExpLast As String
    Dim sTemplate As String
    Dim i As Long

    ResolveSubIdRelocate = -1
    sTemplate = ents(idx).Template
    sExpLast = ExpandTemplate(sTemplate, ents(idx).ValNames, ents(idx).ValValues, ents(idx).nVals)

    For i = 0 To nBearers - 1
        If Not IsSubIdClaimed(claimed, i) Then
            If texts(i) = sExpLast Then
                ResolveSubIdRelocate = ClaimSubId(claimed, i)
                Exit Function
            End If
        End If
    Next i
    For i = 0 To nBearers - 1
        If Not IsSubIdClaimed(claimed, i) Then
            If texts(i) = sTemplate Then
                ResolveSubIdRelocate = ClaimSubId(claimed, i)
                Exit Function
            End If
        End If
    Next i

    If ents(idx).SubId < 0 Then Exit Function
    If ents(idx).SubId > nBearers - 1 Then Exit Function
    If IsSubIdClaimed(claimed, ents(idx).SubId) Then Exit Function
    ResolveSubIdRelocate = ClaimSubId(claimed, ents(idx).SubId)
    Exit Function

ErrorHandler:
    ResolveSubIdRelocate = -1
End Function

' Mark a SubId as taken by the entry currently being resolved, and hand it back so the caller reads as
' one statement. The silent handler stands in for a bounds test: claimed() is sized to nBearers by
' ResolveAllSubIds, and an out-of-range or unallocated access must simply not claim anything.
Private Function ClaimSubId(ByRef claimed() As Boolean, ByVal SubId As Long) As Long
    On Error Resume Next
    ClaimSubId = SubId
    claimed(SubId) = True
End Function

Private Function IsSubIdClaimed(ByRef claimed() As Boolean, ByVal SubId As Long) As Boolean
    On Error Resume Next
    IsSubIdClaimed = False
    IsSubIdClaimed = claimed(SubId)
End Function

' Defensive bearer guard (the Ouroboros exclusion): the bearer is always a TOP-LEVEL model element, never a
' cell component. Refuses SILENTLY on purpose. Full rationale: see "Bearer guards" in
' property-rendering-mechanics.md.
Private Function IsAcceptableBearerElement(ByVal oEl As element) As Boolean
    On Error GoTo ErrorHandler

    IsAcceptableBearerElement = False
    If oEl Is Nothing Then Exit Function
    If oEl.IsComponentElement Then Exit Function
    IsAcceptableBearerElement = True
    Exit Function

ErrorHandler:
    IsAcceptableBearerElement = False
End Function

' v1 refuses to author a token inside a TEXTNODE belonging to a cell fed by an active group source: the
' exclusion granularity is one whole bearer, and a TextNode IS one bearer. Full rationale: see "Bearer
' guards" in property-rendering-mechanics.md.
Private Function CanBearTokens(ByVal oEl As element, ByVal sBearerText As String) As Boolean
    On Error GoTo ErrorHandler

    CanBearTokens = True
    If Not oEl.IsCellElement Then Exit Function
    If Not PropertyCalculation.IsTriggerCell(oEl) Then Exit Function

    ' A multi-line bearer is a TextNode: the whole-text form joins its lines with vbLf.
    If InStr(1, sBearerText, vbLf) > 0 Then CanBearTokens = False
    Exit Function

ErrorHandler:
    CanBearTokens = False
End Function

' Bind-time discoverability: tell the user which engine owns the value they are about to display (manual
' edits of a rendered value ARE overwritten, by design), and warn about the one static cycle v1 can
' detect - a token rendered inside the very cell whose text feeds the property through a CellText rule.
Private Sub WarnGovernedValue(ByVal oEl As element, ByVal P As String)
    On Error Resume Next

    Dim kind As CalcSource
    Dim sArg As String
    Dim sCanonical As String

    If Not PropertyCalculation.GetCalcRuleForProperty(P, oEl, kind, sArg, sCanonical) Then Exit Sub

    If Not mbGovernedShown Then
        ShowStatus LangManager.GetTranslation("RenderValueGoverned", P)
        mbGovernedShown = True
    End If

    If kind = csCellText Then
        If oEl.IsCellElement Then ReportCycleWarning
    End If
End Sub

' "This value is computed from the text of this very cell; the rendered text is excluded from that
' computation." Used both at bind time, as the static-cycle warning, and by the top-up scan when it
' refuses the cell's LAST unrendered sub-text - in both cases the message is exactly the point being made.
Private Sub ReportCycleWarning()
    On Error Resume Next
    If Not mbCycleShown Then
        LangManager.ShowStatusT "RenderCycleWarning"
        mbCycleShown = True
    End If
End Sub

'######################################################################################################################
'                                   METADATA (ARES_SYS / ARES_Render)
'######################################################################################################################

' Read and parse the ARES_Render metadata of an element. False = attached but unusable - the caller must
' refuse, never guess. True with nEnts = 0 is the legal "freshly attached, nothing stored yet" state. EVERY
' access names ItemName AND LibraryName explicitly (two independent CustomPropertyHandler traps otherwise).
' Full rationale: see "ReadRenderMetadata" in property-rendering-mechanics.md.
Private Function ReadRenderMetadata(ByVal El As element, ByRef ents() As RenderEntry, ByRef nEnts As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim vSchema As Variant
    Dim vEntries As Variant
    Dim sSchema As String
    Dim sEntries As String

    ReadRenderMetadata = False
    nEnts = 0

    vSchema = CustomPropertyHandler.GetPropertyValueFromElement(El, PROP_SCHEMA, ARES_ITEM_RENDER, ARES_NAME_LIBRARY_SYS, True)
    vEntries = CustomPropertyHandler.GetPropertyValueFromElement(El, PROP_ENTRIES, ARES_ITEM_RENDER, ARES_NAME_LIBRARY_SYS, True)

    sSchema = VariantToPlainString(vSchema)
    sEntries = VariantToPlainString(vEntries)

    If Len(sSchema) = 0 Then
        ' Freshly attached: both fields empty is legal and means "nothing stored yet". A version-less
        ' item that nevertheless carries entries has been tampered with.
        If Len(sEntries) > 0 Then
            ReportMetadataUnreadable
            Exit Function
        End If
        ReadRenderMetadata = True
        Exit Function
    End If

    If sSchema <> ARES_RENDER_SCHEMA Then
        ' A newer ARES wrote a shape this build cannot read. Refuse to interpret it and NEVER rewrite it.
        ReportSchemaUnsupported
        Exit Function
    End If

    If Len(sEntries) = 0 Then
        ReadRenderMetadata = True
        Exit Function
    End If

    ReadRenderMetadata = DeserializeEntries(sEntries, ents, nEnts)
    If Not ReadRenderMetadata Then ReportMetadataUnreadable
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.ReadRenderMetadata"
    ReportMetadataUnreadable
    ReadRenderMetadata = False
    nEnts = 0
End Function

' Write SchemaVersion + Entries back. Both writes are strict (bNoFallback) and both name their item and
' library. Called only AFTER the corresponding text writes succeeded.
Private Function WriteRenderMetadata(ByVal El As element, ByRef ents() As RenderEntry, ByVal nEnts As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim sEntries As String

    WriteRenderMetadata = False
    sEntries = SerializeEntries(ents, nEnts)

    If Not CustomPropertyHandler.SetPropertyValueToElement(El, PROP_SCHEMA, ARES_RENDER_SCHEMA, ARES_ITEM_RENDER, ARES_NAME_LIBRARY_SYS, True) Then Exit Function
    If Not CustomPropertyHandler.SetPropertyValueToElement(El, PROP_ENTRIES, sEntries, ARES_ITEM_RENDER, ARES_NAME_LIBRARY_SYS, True) Then Exit Function

    ' No Rewrite on the bearer: an item write goes straight to the file (mvba-docs/03-methods/
    ' SetPropertyValue_Method.md - "always writes a change back to the file immediately"). The targeted
    ' text write does its own Rewrite on the sub-element it touched.
    WriteRenderMetadata = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.WriteRenderMetadata"
    WriteRenderMetadata = False
End Function

' Records separated by Chr(1), fields by Chr(2), LastValues by Chr(3) (a flat name,value,name,value
' sequence), line breaks inside a Template by Chr(4). These four cannot be typed into a MicroStation text
' nor produced by a normal property value, which buys one fail-closed rejection rule instead of a whole
' escaping machinery: a VALUE carrying any of them (or CR/LF) is refused at the render choke point.
Private Function SerializeEntries(ByRef ents() As RenderEntry, ByVal nEnts As Long) As String
    On Error GoTo ErrorHandler

    Dim sOut As String
    Dim sRec As String
    Dim i As Long, j As Long

    SerializeEntries = ""
    If nEnts = 0 Then Exit Function

    For i = 0 To nEnts - 1
        If Not ents(i).Dropped Then
            sRec = CStr(ents(i).SubId) & SepField() & EncodeTemplate(ents(i).Template) & SepField()
            For j = 0 To ents(i).nVals - 1
                If j > 0 Then sRec = sRec & SepPair()
                sRec = sRec & ents(i).ValNames(j) & SepPair() & ents(i).ValValues(j)
            Next j
            If Len(sOut) > 0 Then sOut = sOut & SepRecord()
            sOut = sOut & sRec
        End If
    Next i

    SerializeEntries = sOut
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.SerializeEntries"
    SerializeEntries = ""
End Function

' Parse the Entries blob. False on ANY malformation - hostile input is never partially trusted.
Private Function DeserializeEntries(ByVal sBlob As String, ByRef ents() As RenderEntry, ByRef nEnts As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim recs() As String
    Dim fields() As String
    Dim pairs() As String
    Dim names() As String
    Dim values() As String
    Dim n As Long
    Dim i As Long, j As Long

    DeserializeEntries = False
    nEnts = 0

    recs = Split(sBlob, SepRecord())
    ' A zero-length blob yields UBound = -1; iterating LBound..UBound is the only safe form (never
    ' index parts(0) directly).
    For i = LBound(recs) To UBound(recs)
        If Len(recs(i)) > 0 Then
            fields = Split(recs(i), SepField())
            If UBound(fields) - LBound(fields) <> 2 Then Exit Function

            If Not IsNumeric(fields(LBound(fields))) Then Exit Function

            n = 0
            If Len(fields(LBound(fields) + 2)) > 0 Then
                pairs = Split(fields(LBound(fields) + 2), SepPair())
                If ((UBound(pairs) - LBound(pairs) + 1) Mod 2) <> 0 Then Exit Function
                For j = LBound(pairs) To UBound(pairs) Step 2
                    AppendValue pairs(j), pairs(j + 1), names, values, n
                Next j
            End If

            ReDim Preserve ents(0 To nEnts)
            ents(nEnts).SubId = CLng(fields(LBound(fields)))
            ents(nEnts).Template = DecodeTemplate(fields(LBound(fields) + 1))
            ents(nEnts).Dropped = False
            SetEntryValues ents, nEnts, names, values, n
            nEnts = nEnts + 1
        End If
    Next i

    DeserializeEntries = True
    Exit Function

ErrorHandler:
    ' Malformed metadata is expected hostile input, not a fault worth logging: the caller surfaces a
    ' status and refuses to render.
    DeserializeEntries = False
    nEnts = 0
End Function

' A Template legitimately carries line breaks (a multi-line TextNode's Template is exactly what
' GetTextAtSubId joins with vbLf), so they are serialised rather than refused - the refusal rule applies
' to VALUES only.
Private Function EncodeTemplate(ByVal s As String) As String
    EncodeTemplate = Replace(s, vbLf, SepLine())
End Function

Private Function DecodeTemplate(ByVal s As String) As String
    DecodeTemplate = Replace(s, SepLine(), vbLf)
End Function

' VBA forbids a function call in a Const initialiser, so the four delimiters are trivial functions.
Private Function SepRecord() As String
    SepRecord = Chr(1)
End Function

Private Function SepField() As String
    SepField = Chr(2)
End Function

Private Function SepPair() As String
    SepPair = Chr(3)
End Function

Private Function SepLine() As String
    SepLine = Chr(4)
End Function

' True when a string carries one of the four serialisation delimiters. None of them can be typed into a
' MicroStation text nor produced by a normal property value, which is exactly why they were chosen.
Private Function ContainsSerialisationDelimiter(ByVal s As String) As Boolean
    ContainsSerialisationDelimiter = True
    If InStr(1, s, SepRecord()) > 0 Then Exit Function
    If InStr(1, s, SepField()) > 0 Then Exit Function
    If InStr(1, s, SepPair()) > 0 Then Exit Function
    If InStr(1, s, SepLine()) > 0 Then Exit Function
    ContainsSerialisationDelimiter = False
End Function

' The strict rule for VALUES: no delimiter and no line break of any kind. A Template is deliberately NOT
' held to it - a multi-line TextNode's Template legitimately contains vbLf, which is what SepLine
' serialises.
Private Function ValueHasIllegalChar(ByVal s As String) As Boolean
    ValueHasIllegalChar = True
    If ContainsSerialisationDelimiter(s) Then Exit Function
    If InStr(1, s, vbCr) > 0 Then Exit Function
    If InStr(1, s, vbLf) > 0 Then Exit Function
    ValueHasIllegalChar = False
End Function

'######################################################################################################################
'                                   TEMPLATE PARSING / VALUE MAP HELPERS
'######################################################################################################################

' Split a Template into its alternating literals and tokens: lits(0..nTok) and toks(0..nTok-1).
' bValidateNames = True (the runtime) accepts only tokens naming a property the DGNLib actually declares;
' anything else stays part of the surrounding literal, fail-closed. False gives the pure grammar, which
' is what the unit tests assert without MicroStation.
Private Function ParseTemplate(ByVal sTemplate As String, ByRef lits() As String, ByRef toks() As String, ByRef nTok As Long, ByVal bValidateNames As Boolean) As Boolean
    On Error GoTo ErrorHandler

    Dim cursor As Long
    Dim posOpen As Long
    Dim posClose As Long
    Dim sName As String
    Dim sLit As String

    ParseTemplate = False
    nTok = 0
    ReDim lits(0 To 0)
    lits(0) = ""
    cursor = 1
    sLit = ""

    Do
        posOpen = InStr(cursor, sTemplate, TOKEN_OPEN, vbTextCompare)
        If posOpen = 0 Then Exit Do

        posClose = InStr(posOpen + Len(TOKEN_OPEN), sTemplate, TOKEN_CLOSE)
        If posClose = 0 Then Exit Do

        sName = Mid(sTemplate, posOpen + Len(TOKEN_OPEN), posClose - posOpen - Len(TOKEN_OPEN))

        If IsAcceptableTokenName(sName, bValidateNames) Then
            sLit = sLit & Mid(sTemplate, cursor, posOpen - cursor)
            lits(nTok) = sLit
            ReDim Preserve toks(0 To nTok)
            toks(nTok) = sName
            nTok = nTok + 1
            ReDim Preserve lits(0 To nTok)
            lits(nTok) = ""
            sLit = ""
            cursor = posClose + Len(TOKEN_CLOSE)
        Else
            ' Unknown / malformed name: the whole "Prop[...]" run stays literal text.
            sLit = sLit & Mid(sTemplate, cursor, posClose + Len(TOKEN_CLOSE) - cursor)
            cursor = posClose + Len(TOKEN_CLOSE)
            If bValidateNames Then ReportTokenUnknown
        End If
    Loop

    lits(nTok) = sLit & Mid(sTemplate, cursor)
    ParseTemplate = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.ParseTemplate"
    ParseTemplate = False
    nTok = 0
End Function

' A Template is well formed when no property is tokenised twice (v1: ONE token per property per text) and
' no two tokens are adjacent - adjacency produces an interior EMPTY literal, leaving the span boundary
' undefined for the per-token release. Public test seam, like ExpandTemplate/AlignVisible.
Public Function TemplateIsWellFormed(ByVal sTemplate As String, Optional ByVal bValidateNames As Boolean = True) As Boolean
    On Error GoTo ErrorHandler

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim i As Long, j As Long

    TemplateIsWellFormed = False
    If Not ParseTemplate(sTemplate, lits, toks, nTok, bValidateNames) Then Exit Function
    If nTok = 0 Then
        TemplateIsWellFormed = True
        Exit Function
    End If

    For i = 0 To nTok - 1
        For j = i + 1 To nTok - 1
            If StrComp(toks(i), toks(j), vbTextCompare) = 0 Then Exit Function
        Next j
    Next i

    For i = 1 To nTok - 1
        If Len(lits(i)) = 0 Then Exit Function
    Next i

    TemplateIsWellFormed = True
    Exit Function

ErrorHandler:
    TemplateIsWellFormed = False
End Function

' Report the exact structural reason a first author was refused (duplicate vs adjacent tokens).
Private Sub ReportStructuralRefusal(ByVal sTemplate As String)
    On Error Resume Next

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim i As Long, j As Long

    If Not ParseTemplate(sTemplate, lits, toks, nTok, True) Then Exit Sub

    For i = 0 To nTok - 1
        For j = i + 1 To nTok - 1
            If StrComp(toks(i), toks(j), vbTextCompare) = 0 Then
                ReportDuplicateToken
                Exit Sub
            End If
        Next j
    Next i

    ReportAdjacentTokens
End Sub

' The exact substring a token occupies in the text. This is the "unset" cue Expand writes, and the thing
' the conservative fallback keeps.
Private Function TokenLiteral(ByVal sName As String) As String
    TokenLiteral = TOKEN_OPEN & sName & TOKEN_CLOSE
End Function

' A token name is acceptable when it is non-empty, carries no grammar metacharacter, and - at runtime -
' names a property the DGNLib really declares.
Private Function IsAcceptableTokenName(ByVal sName As String, ByVal bValidateNames As Boolean) As Boolean
    IsAcceptableTokenName = False
    If Len(Trim(sName)) = 0 Then Exit Function
    If InStr(1, sName, "[") > 0 Then Exit Function
    If InStr(1, sName, ";") > 0 Then Exit Function
    If ValueHasIllegalChar(sName) Then Exit Function
    If Not bValidateNames Then
        IsAcceptableTokenName = True
        Exit Function
    End If
    IsAcceptableTokenName = IsKnownProperty(sName)
End Function

' Membership test against the cached DGNLib property names. Matching is case-insensitive, aligned with
' the rest of the system.
Private Function IsKnownProperty(ByVal sName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim i As Long

    IsKnownProperty = False
    EnsurePropertyNames
    For i = 0 To mnPropNames - 1
        If StrComp(msPropNames(i), sName, vbTextCompare) = 0 Then
            IsKnownProperty = True
            Exit Function
        End If
    Next i
    Exit Function

ErrorHandler:
    IsKnownProperty = False
End Function

' Populate the property-name cache, re-reading it whenever the active design file changed. The [""]
' sentinel GetCustomPropertyNames returns for an absent/empty library is dropped, so an empty name can
' never validate a token.
Private Sub EnsurePropertyNames()
    On Error GoTo ErrorHandler

    Dim sFile As String
    Dim names() As String
    Dim i As Long

    sFile = ActiveDesignFileName()
    If mbNamesCached Then
        If sFile = msCachedFor Then Exit Sub
    End If

    mnPropNames = 0
    names = CustomPropertyHandler.GetCustomPropertyNames()
    For i = LBound(names) To UBound(names)
        If Len(Trim(names(i))) > 0 Then
            ReDim Preserve msPropNames(0 To mnPropNames)
            msPropNames(mnPropNames) = names(i)
            mnPropNames = mnPropNames + 1
        End If
    Next i

    msCachedFor = sFile
    mbNamesCached = True
    ' The ARES_SYS probe shares the same invalidation point.
    mbSysChecked = False
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.EnsurePropertyNames"
    mnPropNames = 0
    mbNamesCached = False
End Sub

' Is the internal ARES_SYS library deployed? Cached with the property names; a station whose DGNLib
' predates epic 15 self-disables the whole feature rather than half-write anything.
Private Function IsSysLibraryPresent() As Boolean
    On Error GoTo ErrorHandler

    IsSysLibraryPresent = False
    EnsurePropertyNames
    If Not mbSysChecked Then
        mbSysPresent = Not (CustomPropertyHandler.FindItemTypeLibrary(ARES_NAME_LIBRARY_SYS) Is Nothing)
        mbSysChecked = True
    End If
    IsSysLibraryPresent = mbSysPresent
    Exit Function

ErrorHandler:
    IsSysLibraryPresent = False
End Function

' Active design file identity, used only as the cache invalidation key. Silent by design.
Private Function ActiveDesignFileName() As String
    On Error Resume Next
    ActiveDesignFileName = ""
    ActiveDesignFileName = ActiveDesignFile.FullName
End Function

' Look up a token's value in a name/value pair set (case-insensitive). Absent or empty both yield "",
' which Expand renders as the literal token - the two cases are deliberately indistinguishable there.
Private Function LookupValue(ByVal sName As String, ByRef names() As String, ByRef values() As String, ByVal n As Long) As String
    Dim i As Long
    LookupValue = ""
    If n <= 0 Then Exit Function
    For i = 0 To n - 1
        If StrComp(names(i), sName, vbTextCompare) = 0 Then
            LookupValue = values(i)
            Exit Function
        End If
    Next i
End Function

' Does an ENTRY exist for this token (regardless of whether its value is empty)? The unset case turns on
' this distinction: an existing-but-empty entry means "rendered as the literal token last time".
Private Function HasValueEntry(ByVal sName As String, ByRef names() As String, ByVal n As Long) As Boolean
    Dim i As Long
    HasValueEntry = False
    If n <= 0 Then Exit Function
    For i = 0 To n - 1
        If StrComp(names(i), sName, vbTextCompare) = 0 Then
            HasValueEntry = True
            Exit Function
        End If
    Next i
End Function

' Append one name/value pair, growing both parallel arrays together.
Private Sub AppendValue(ByVal sName As String, ByVal sValue As String, ByRef names() As String, ByRef values() As String, ByRef n As Long)
    ReDim Preserve names(0 To n)
    ReDim Preserve values(0 To n)
    names(n) = sName
    values(n) = sValue
    n = n + 1
End Sub

' Replace an entry's stored LastValues wholesale.
Private Sub SetEntryValues(ByRef ents() As RenderEntry, ByVal idx As Long, ByRef names() As String, ByRef values() As String, ByVal n As Long)
    Dim i As Long
    ents(idx).nVals = n
    If n <= 0 Then Exit Sub
    ReDim ents(idx).ValNames(0 To n - 1)
    ReDim ents(idx).ValValues(0 To n - 1)
    For i = 0 To n - 1
        ents(idx).ValNames(i) = names(i)
        ents(idx).ValValues(i) = values(i)
    Next i
End Sub

' Validation on read: the Template's token set and the LastValues key set must be the SAME set. Anything
' else is metadata that was hand-edited through the native Properties pane, and it is never rendered as
' if it had been intended.
Private Function EntryIsConsistent(ByRef ent As RenderEntry) As Boolean
    On Error GoTo ErrorHandler

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim i As Long

    EntryIsConsistent = False
    If Not ParseTemplate(ent.Template, lits, toks, nTok, True) Then Exit Function
    If nTok <> ent.nVals Then Exit Function

    For i = 0 To nTok - 1
        If Not HasValueEntry(toks(i), ent.ValNames, ent.nVals) Then Exit Function
    Next i

    EntryIsConsistent = TemplateIsWellFormed(ent.Template, True)
    Exit Function

ErrorHandler:
    EntryIsConsistent = False
End Function

'######################################################################################################################
'                                          SMALL HELPERS
'######################################################################################################################

' v1 scope: Text / TextNode / Cell only, active model only. Everything else keeps its tokens literal.
Private Function IsRenderableType(ByVal El As element) As Boolean
    On Error GoTo ErrorHandler

    IsRenderableType = False
    If El Is Nothing Then Exit Function
    Select Case True
        Case El.IsTextElement
            IsRenderableType = True
        Case El.IsTextNodeElement
            IsRenderableType = True
        Case El.IsCellElement
            IsRenderableType = True
    End Select
    Exit Function

ErrorHandler:
    IsRenderableType = False
End Function

' Convert an item read to a plain String WITHOUT any locale-dependent transform: Null and any non-string
' type yield "". Used for the metadata's own String fields.
Private Function VariantToPlainString(ByVal v As Variant) As String
    VariantToPlainString = ""
    If IsNull(v) Then Exit Function
    If IsArray(v) Then Exit Function
    If VarType(v) <> vbString Then Exit Function
    VariantToPlainString = v
End Function

' Have we already walked this graphic group in the current drain? Records it if not.
Private Function NoteGroupSeen(ByVal lGroup As Long, ByRef seen() As Long, ByRef nSeen As Long) As Boolean
    Dim i As Long
    NoteGroupSeen = False
    For i = 0 To nSeen - 1
        If seen(i) = lGroup Then
            NoteGroupSeen = True
            Exit Function
        End If
    Next i
    ReDim Preserve seen(0 To nSeen)
    seen(nSeen) = lGroup
    nSeen = nSeen + 1
End Function

' How many entries would actually be written back (a released entry no longer counts).
Private Function CountLiveEntries(ByRef ents() As RenderEntry, ByVal nEnts As Long) As Long
    Dim i As Long
    CountLiveEntries = 0
    For i = 0 To nEnts - 1
        If Not ents(i).Dropped Then CountLiveEntries = CountLiveEntries + 1
    Next i
End Function

' Safe "array has at least one element" check (mirrors PropertyCalculation.HasElements). UBound returns
' -1 for an empty array and RAISES for an uninitialised one.
Private Function HasElements(ByRef arr() As element) As Boolean
    On Error Resume Next
    HasElements = False
    If UBound(arr) <> -1 Then HasElements = True
    On Error GoTo 0
End Function

'######################################################################################################################
'                              ONE-SHOT STATUS REPORTERS (status bar only, translated)
'######################################################################################################################

' Every refusal below is EXPECTED user feedback, so it is status-bar only and never logged - an expected
' refusal in the .log is a project rule violation. The two schema/library self-disable conditions are the
' single mandated exception: they also write ONE English line, because they mean the feature turned
' itself off and that must be greppable.
Private Sub ResetOneShots()
    On Error Resume Next
    mbTokenUnknownShown = False
    mbValueUnsupportedShown = False
    mbValueIllegalShown = False
    mbMetadataInvalidShown = False
    mbMetadataUnreadableShown = False
    mbSchemaShown = False
    mbLibraryMissingShown = False
    mbAmbiguousShown = False
    mbBindingReleasedShown = False
    mbLockedShown = False
    mbDriftShown = False
    mbDuplicateShown = False
    mbAdjacentShown = False
    mbTextNodeRefusedShown = False
    mbNotBoundShown = False
    mbGovernedShown = False
    mbCycleShown = False
End Sub

Private Sub ReportTokenUnknown()
    On Error Resume Next
    If Not mbTokenUnknownShown Then
        LangManager.ShowStatusT "RenderTokenUnknown"
        mbTokenUnknownShown = True
    End If
End Sub

Private Sub ReportValueUnsupported()
    On Error Resume Next
    If Not mbValueUnsupportedShown Then
        LangManager.ShowStatusT "RenderValueUnsupported"
        mbValueUnsupportedShown = True
    End If
End Sub

Private Sub ReportValueIllegal()
    On Error Resume Next
    If Not mbValueIllegalShown Then
        LangManager.ShowStatusT "RenderValueIllegalChars"
        mbValueIllegalShown = True
    End If
End Sub

Private Sub ReportMetadataInvalid()
    On Error Resume Next
    If Not mbMetadataInvalidShown Then
        LangManager.ShowStatusT "RenderMetadataInvalid"
        mbMetadataInvalidShown = True
    End If
End Sub

Private Sub ReportMetadataUnreadable()
    On Error Resume Next
    If Not mbMetadataUnreadableShown Then
        LangManager.ShowStatusT "RenderMetadataUnreadable"
        mbMetadataUnreadableShown = True
    End If
End Sub

' Self-disable condition: status AND ONE English log line - the single mandated exception to the
' "expected refusals are never logged" rule, because a feature that turned itself off must be greppable.
' The log is gated by the same one-shot flag so a whole batch cannot flood the .log.
Private Sub ReportSchemaUnsupported()
    On Error Resume Next
    If Not mbSchemaLogged Then
        ErrorHandler.HandleError "Property rendering: unsupported ARES_Render schema version, metadata left untouched", 0, "", "PropertyRendering.ReadRenderMetadata"
        mbSchemaLogged = True
    End If
    If Not mbSchemaShown Then
        LangManager.ShowStatusT "RenderSchemaUnsupported"
        mbSchemaShown = True
    End If
End Sub

' Self-disable condition: status AND ONE English log line (see ReportSchemaUnsupported).
Private Sub ReportLibraryMissing()
    On Error Resume Next
    If Not mbLibraryLogged Then
        ErrorHandler.HandleError "Property rendering: internal ARES_SYS item type library not found, rendering disabled", 0, "", "PropertyRendering.IsSysLibraryPresent"
        mbLibraryLogged = True
    End If
    If Not mbLibraryMissingShown Then
        LangManager.ShowStatusT "RenderLibraryMissing"
        mbLibraryMissingShown = True
    End If
End Sub

' Third self-disable condition, and the one that is not a refusal but a FAULT LOOP guard. Full rationale:
' see "Session-scoped self-disable guards" in property-rendering-mechanics.md.
Private Sub DisableAfterWriteFailure()
    On Error Resume Next
    If mbWriteDisabled Then Exit Sub
    mbWriteDisabled = True
    ErrorHandler.HandleError "Property rendering: ARES_Render metadata write failed, rendering disabled for this session", 0, "", "PropertyRendering.WriteRenderMetadata"
End Sub

Private Sub ReportAmbiguous()
    On Error Resume Next
    If Not mbAmbiguousShown Then
        LangManager.ShowStatusT "RenderAmbiguousEdit"
        mbAmbiguousShown = True
    End If
End Sub

Private Sub ReportBindingReleased()
    On Error Resume Next
    If Not mbBindingReleasedShown Then
        LangManager.ShowStatusT "RenderBindingReleased"
        mbBindingReleasedShown = True
    End If
End Sub

Private Sub ReportLocked()
    On Error Resume Next
    If Not mbLockedShown Then
        LangManager.ShowStatusT "RenderElementLocked"
        mbLockedShown = True
    End If
End Sub

Private Sub ReportDrift()
    On Error Resume Next
    If Not mbDriftShown Then
        LangManager.ShowStatusT "RenderSubIdDrift"
        mbDriftShown = True
    End If
End Sub

Private Sub ReportDuplicateToken()
    On Error Resume Next
    If Not mbDuplicateShown Then
        LangManager.ShowStatusT "RenderDuplicateToken"
        mbDuplicateShown = True
    End If
End Sub

Private Sub ReportAdjacentTokens()
    On Error Resume Next
    If Not mbAdjacentShown Then
        LangManager.ShowStatusT "RenderAdjacentTokens"
        mbAdjacentShown = True
    End If
End Sub

Private Sub ReportTextNodeRefused()
    On Error Resume Next
    If Not mbTextNodeRefusedShown Then
        LangManager.ShowStatusT "RenderTextNodeInCellRefused"
        mbTextNodeRefusedShown = True
    End If
End Sub

Private Sub ReportNotBound()
    On Error Resume Next
    If Not mbNotBoundShown Then
        LangManager.ShowStatusT "RenderNotBound"
        mbNotBoundShown = True
    End If
End Sub
