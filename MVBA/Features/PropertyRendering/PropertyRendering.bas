' Module: PropertyRendering
' Description: The SOLE text writer of ARES - the third engine, after Tagging (the sole attacher) and
'              Calculation (the sole value writer). It displays a custom property's value inside a text
'              by FULLY REPLACING a "Prop[Name]" token typed in that text: the user writes
'              "Ligne Prop[Len] m", the renderer writes "Ligne 13,3 m". Rendered text is CLEAN - no
'              visible scaffolding ever survives a render.
'
'              The binding is NOT the visible text and NOT the graphic group: it is hidden metadata, an
'              ItemType ARES_Render living in a SECOND, internal ItemTypeLibrary (ARES_SYS) authored in
'              the same ARES_Custom_Properties.dgnlib. ARES_Render carries two String properties:
'              SchemaVersion, and Entries - a serialised list of {SubId, Template, LastValues} records.
'              LastValues maps each token's property name to the value last rendered for it. The last
'              rendering is never stored: it is COMPUTED as Expand(Template, LastValues).
'
'              That computed last rendering is what makes a hidden template safe. On each processed
'              bearer, per entry, after an IsLocked test and a successful non-empty read:
'                1. visible = Expand(Template, CURRENT values)  -> up to date, strict no-op (terminator)
'                2. visible = Expand(Template, LastValues), or visible IS the Template up to CASE (an
'                   entry authored but never rendered) -> re-render + store
'                3. otherwise (a NON-EMPTY read differing from BOTH) -> the user edited the text:
'                   PER-TOKEN release. A rendered span left intact is substituted back to its token in
'                   the NEW Template and keeps updating; a span the user changed drops its token and
'                   becomes static text. Ambiguous alignment -> conservative fallback + status.
'                4. no metadata at all, visible carries valid tokens -> FIRST AUTHOR, hybrid policy:
'                   auto-bind only when every token's property is ALREADY attached to the element
'                   (attachment is the intent signal); otherwise the token stays literal and the
'                   BindPropertyRender key-in is the entry point.
'              An EMPTY/absent value renders the token's own literal text as the visible "unset" cue AND
'              stores an EMPTY LastValues entry - without that write, an unset -> reset round-trip lands
'              in branch 3 and silently flattens the binding.
'
'              DOCTRINE: this module writes VISIBLE TEXT and its own ARES_SYS metadata, nothing else. It
'              never writes a USER property value (that is Calculation's job) and never attaches or
'              detaches anything itself - the ARES_Render attach/detach is delegated to the two
'              PropertyTagging wrappers so the attach choke point stays unique. Any
'              SetPropertyValueToElement on the "ARES" library inside this module is a review BLOCKER.
'
'              Values are copied VERBATIM (decision: convert once at value-write, never at render) - no
'              CStr, no Format, no locale-dependent transform, so two stations never rewrite each other's
'              decimal separator. All dressing is static template text: "(Prop[Len]m)" renders "(12.3m)",
'              byte-identical to legacy AutoLengths output.
'
'              Coexistence with AutoLengths: on a text carrying a token, THE RENDERER WINS. It is enforced
'              in ONE direction only - ElementChangeHandler's Branch 1 skips any render-bound element via
'              IsRenderBound. The renderer does NOT check whether its own expansion looks like a legacy
'              trigger: it deliberately writes "(12.3m)" when that is what the template says, which is the
'              flagship "(Prop[Len]m)" case and the whole point of superseding AutoLengths. A text the
'              renderer does not own is untouched and stays legacy territory - including one whose binding
'              is later RELEASED, which hands it back to AutoLengths, as it should.
'
'              Every user-facing refusal is status-bar only, translated and one-shot; only the
'              schema/library self-disable conditions also log ONE English line.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (global ARESConfig), CustomPropertyHandler, PropertyTagging,
'               PropertyCalculation, StringsInEl, Link, LangManager, ErrorHandlerClass (global ErrorHandler)

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

' D6 - WHITELIST of the SYMBOLS an accepted addition may border the value with (letters are admitted by
' code range in IsSafeBoundaryChar, digits never are). Deliberately a whitelist, never a blacklist: a
' blacklist has to enumerate every character that could pass for part of a number and it silently loses the
' moment one is missed - Chr(160) (non-breaking space, the French thousands separator Excel emits in fr-FR),
' Chr(8239) (narrow no-break space), the Unicode digits. Every one of those falls OUTSIDE this list and is
' refused without ever being named, which is what makes the "no forged number" guarantee closed rather than
' a list of the cases we happened to think of.
'
' ADMISSION CRITERION, applied one character at a time to what follows: placed between the value and the
' addition, the character must not let the whole be read as ONE number of a DIFFERENT value. That excludes
' strictly MORE than IsNumericText's own alphabet ("0123456789 ,."), and each exclusion has a name:
'   '   Swiss thousands separator - a suffix of "'500" would render a future 20 as "20'500"
'   /   fraction - a suffix of "/2" would render a future 1 as "1/2"
'   - + sign - a PREFIX of "-" would render a future 20 as "-20"
' Those three are excluded WHOLESALE rather than per direction ("135-A" after a terminal token therefore
' releases, as it does today): one list that can be audited character by character is worth more than two
' lists that can drift apart, and refusing costs nothing but today's behaviour.
'
' Two admitted characters are worth naming so the next reader does not have to re-derive them:
'   ( )  accounting negatives need BOTH, and this rule can only ever contribute ONE of them - a
'        simultaneous prefix AND suffix fails the anchor test and releases before reaching this list.
'        The pair CAN still occur, and saying otherwise would be wrong: on a Template AUTHORED as
'        "Prop[Len])", typing "(" in front is accepted here and a future value of 20 renders "(20)".
'        What makes that acceptable is not that it cannot happen - it is that the other half was
'        WRITTEN BY THE USER and is visible in the text. This rule forged nothing; it preserved a
'        binding inside a composition its author already had on screen. The guarantee is about what
'        an ADDITION can weld onto a value, not about what an authored Template can spell.
'   :    "20:30" reads as a time, not as a number of a different value.
' Letters are admitted, but never BETWEEN two digits - see the exponent guard in the two callers, which
' closes "20e5" and "0x20" by shape instead of by naming e and x.
Private Const SAFE_BOUNDARY_SYMBOLS As String = "%()[]{}\*#~<>=:;!?&@_|"""

' D8 - every character that can appear INSIDE a number literal of any base: decimal digits, the hex digits,
' and the base markers. Used only by NumericTailIsPossible, to delimit the trailing run it examines - the
' run's FIRST character is what decides, so this list bounds the scan and never grants safety by itself.
'
' THE LIST CAN BE WRONG WITHOUT THE RULE BEING WRONG, and that is the whole reason this shape was chosen.
' Too broad costs nothing; too narrow only breaks the run EARLIER, which cannot manufacture a hole. It is
' the exact inverse of the round-38 blacklist, where a missing member WAS the hole. Anything added here
' must keep that property: this list may only ever bound a scan, never grant safety.
Private Const NUMBER_CAPABLE_CHARS As String = "0123456789abcdefABCDEFxXbBoO"

' The two self-disable conditions also write an English log line. Those flags are SESSION-scoped, not
' per-element: ResetOneShots must not clear them, or a station whose DGNLib predates epic 15 would log
' one line per processed element for the whole session.
Private mbLibraryLogged As Boolean
Private mbSchemaLogged As Boolean

' Third SESSION-scoped self-disable, same shape and for the same kind of reason. A metadata write that
' FAILS restores the text it wrote and leaves the persisted state byte-identical to what it found, so the
' next pass takes exactly the same path and emits exactly the same Rewrites - each of which re-queues the
' element. That is an unbounded write/restore loop, and it is deterministically reachable on a DGNLib
' whose ARES_Render properties are not named SchemaVersion/Entries, because bNoFallback deliberately
' removes the tolerance that would otherwise paper over it. One English log line, then the engine is inert
' for the rest of the session (RefreshRenderCaches clears it, for the tests).
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

' The Depth-0 hook, called from ElementChangeHandler.ProcessElement right after the Calculation hook and
' BEFORE the graphic-group filter, so a bound text carrying no graphic group is still rendered on its own
' pass. Values are therefore fresh from the same pass: tagging -> calc -> render.
'
' KNOWN v1 LIMIT, accepted rather than worked around: this hook only ever sees elements the pipeline
' QUEUES, and ShouldQueueElement is deliberately untouched (it is the capture hot path). A bound text
' therefore refreshes from a value pushed by ANOTHER element only when it belongs to a REAL graphic group
' - the calc transport and the repaint hop's Link.GetLink walk both need that group to reach it. Outside a
' group, only direct contact with that very text moves it forward: editing it, or a calc rule that writes
' its own property. Documented in the wiki alongside the DWG/V7 caveat.
'
' Guard order is cost-driven: master switch (no COM cost when OFF) -> cheap TYPE filter -> ONE re-fetch
' -> cheap text walk -> ONE IsItemAttachedToElement. The re-fetch sits exactly between the type filter and
' the first TEXT read: the type cannot go stale, so filtering first keeps every line and arc in the queue
' from paying a COM call, while nothing that reads text ever runs on a handle we did not just fetch. The
' final probe costs a COM Items.Refresh and is unavoidable: a bound text is CLEAN by design, so nothing in
' its visible text distinguishes it from any other text.
Public Sub ProcessElement(ByVal oEl As element)
    On Error GoTo ErrorHandler

    ResetOneShots

    If Not IsEnabled Then Exit Sub
    If oEl Is Nothing Then Exit Sub

    ' A metadata write already failed this session: stay inert instead of replaying the same failing
    ' write, and its restore, on every single pass. The list still drains so nothing leaks across batches.
    If mbWriteDisabled Then GoTo DrainAndExit

    ' Cheap TYPE filter FIRST: Text / TextNode / Cell only. Dimensions, notes, tables and tag elements are
    ' out of v1 scope - their tokens stay literal. This runs on the handle as received, deliberately: it
    ' reads only IsTextElement / IsTextNodeElement / IsCellElement, and an element's TYPE cannot go stale.
    ' Keeping it ahead of the re-fetch is what stops every line and arc in the queue paying a COM call.
    If Not IsRenderableType(oEl) Then GoTo DrainAndExit

    ' RE-FETCH BEFORE READING ANY TEXT, and after the cheap filter so only elements that will really be
    ' processed pay for it. The handle handed to this hook is NOT necessarily current: IdleEventHandler
    ' materialises the WHOLE batch up front (ElementInProcesse.GetAllElements, one GetElementById per
    ' queued id, all of them before the first element is processed), and processing an EARLIER element in
    ' that batch can rewrite the text of a LATER one - the repaint hop walks the graphic group and renders
    ' siblings that are themselves queued. The handle then still serves the text it was fetched with,
    ' while the ITEM side is re-read from the file on every access (CustomPropertyHandler calls
    ' El.Items.Refresh before each read, precisely so a same-pass attach is visible). Stale visible + fresh
    ' values is the one combination the state machine cannot survive: it matches NEITHER expansion, which
    ' branch 3 is entitled to read as positive proof of a user edit, and the binding is released on a text
    ' nobody touched. The read has to be true; weakening branch 3 would only hide it.
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
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.ProcessElement"
    ' The hop must never leak across batches, so it drains even on a fault (the drain is re-entrance
    ' guarded and clears the list unconditionally).
    DrainRepaintHop
End Sub

' Coexistence predicate consumed by ElementChangeHandler's Branch 1: AutoLengths must never write into a
' text this engine owns. IsEnabled FIRST so a render-free configuration pays no COM cost at all on the
' default AUTO_LENGTH And UPDATE_LENGTH = True/True path.
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

' Containment seam consumed by PropertyCalculation's CellText source: the SubIds of oCell's sub-texts
' this engine writes, so a rendered value can never feed the CellText[...] value that governs it (the
' ratcheting cycle). Returns False when the cell carries no readable ARES_Render.
'
' Keyed on metadata PRESENCE only - it must NOT consult IsEnabled. If it did, toggling ARES_Text_Render
' would change what GetConcatenatedText returns and therefore change a CellText VALUE: a feature switch
' must never mutate data.
Public Function GetExcludedSubIds(ByVal oCell As element, ByRef ids() As Long, ByRef nIds As Long) As Boolean
    On Error GoTo ErrorHandler

    GetExcludedSubIds = False
    nIds = 0
    If oCell Is Nothing Then Exit Function

    ' The two cheap gates run on the handle as received. Neither can be fooled by staleness: the library
    ' check does not touch the element at all, and the attach probe reads the ITEM side, which
    ' CustomPropertyHandler refreshes from the file on every access. Only a cell that really carries
    ' ARES_Render gets as far as the re-fetch below - which matters, because calc calls this on EVERY
    ' CellText read.
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

    ' The STORED SubId is NOT the answer. An ordinal drift (the cell was redefined, a sub-element added)
    ' makes it designate a different sub-text, and calc runs BEFORE render inside the same Depth-0 pass,
    ' so this seam would always read the pre-relocation value. Excluding the wrong ordinal hides an
    ' AUTHENTIC source text from the CellText value AND lets the rendered one feed it - the exact
    ' ratcheting cycle this seam exists to prevent. Resolve exactly the way the renderer does, so the
    ' exclusion set is precisely the set of sub-texts the renderer would write.
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
' value STATE CHANGE. Disabled = strict no-op, and that matters: Calculation runs independently of the
' render switch, so without this guard the list would grow for a whole session with nothing draining it.
'
' The SOURCE ELEMENT is stored, not just its group id - Link.GetLink deliberately EXCLUDES the element it
' is given, and that excluded element is precisely the sibling whose value was just written. The group is
' resolved by the renderer at drain time, behind its own IsGraphical guard: letting the caller read
' .GraphicGroup would RAISE on a non-graphical element and pollute ApplyValueToSibling's handler on
' every write.
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

' Drain the repaint hop: for each noted element, run the FULL state machine on the element ITSELF and
' then on its graphic-group siblings, ONE Link.GetLink per distinct group. Never a blind re-render - the
' queue-order race means a sibling may carry a not-yet-processed user edit, and a blind write would
' destroy it. Re-entrance-guarded; the list is cleared unconditionally, so nothing leaks across batches.
'
' TWO RULES HERE ARE LOAD-BEARING, and a group holding TWO bound texts breaks without either of them:
'
'   * Every element is re-fetched from the model immediately before its state machine runs (FreshHandle).
'     The handles this list carries were captured at VALUE-WRITE time, long before anything was rendered,
'     and an element handle does NOT track writes made through another handle - it keeps serving the text
'     it was captured with. Reading a pre-write text is indistinguishable from a user edit, so branch 3
'     fires on a text nobody touched and RELEASES a perfectly healthy binding.
'   * Each element runs AT MOST ONCE per drain (RenderDrainTarget's done-list). With two bound texts in
'     one group, the second one is reached twice - once as another source's Link.GetLink sibling, once as
'     its own noted source - and that second run is exactly the stale read above.
'
' The sibling list is also built BEFORE the source is rendered, because Link.GetLink scans the model off
' the source's own handle and rendering it first would leave that scan running off a handle its own
' Rewrite has just invalidated.
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

    ' Self-disabled after a failed metadata write. This is a REFUSAL the user asked for, so it must be
    ' visible - a silent Exit here would reintroduce exactly the defect ResetOneShots just fixed, and
    ' would make Command.bas's "BindElement reports every refusal" comment false again. The existing
    ' RenderMetadataUnreadable key says it accurately ("the stored binding could not be read or saved"),
    ' so no key is invented for it.
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

' Expand a Template against a set of values. A Template is the alternating sequence L0 T0 L1 T1 ... Ln;
' each token is replaced by its value, and an EMPTY or ABSENT value is replaced by the token's OWN
' LITERAL TEXT ("Prop[Len]"). That literal is the visible "unset" cue, and it is also what makes the
' release-on-empty failure branch unreachable: an emptied value re-materialises its token, so the "no
' token visible" state can never be reached through a value transition.
'
' Public and read-only so UnitTesting can assert the pure logic without MicroStation (VBA cannot call a
' Private proc across modules). bValidateNames:=False makes every "Prop[x]" a token, which is the pure
' grammar; the runtime passes True so an unknown property name stays literal, fail-closed.
'
' bOk reports whether the expansion is TRUSTWORTHY. It matters because this function fails OPEN - it
' returns sTemplate on a fault - and ParseTemplate reaches COM (IsKnownProperty -> EnsurePropertyNames ->
' GetCustomPropertyNames). A caller that compared the visible against that garbage would miss branches 1
' and 2 and fall into branch 3, the only destructive one, turning a transient library-read fault into a
' permanent binding release. The state machine therefore SKIPS the entry when bOk comes back False.
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

' Align a user-edited visible string against the Template that produced the last rendering, and derive
' the NEW Template plus the surviving LastValues (the per-token release).
'
' The literals are walked left to right, each located at or after the current cursor. L0 may legitimately
' be empty (the Template opens with a token) and so may Ln (it ends with one - the final span then runs
' to end-of-string); an INTERIOR empty literal is impossible because adjacent tokens are refused at
' author time, which is exactly what keeps every span boundary defined.
'
' A span SURVIVED when it equals its LastValues entry, or - the unset case - when it equals the token's
' own literal text while that entry is empty. A survivor is substituted back to its token in the new
' Template and keeps updating; anything else was edited, so it becomes static text and its entry is
' dropped. Returns False when a literal cannot be located: the caller must then fall back conservatively
' rather than guess.
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

    ' L0 must sit at the very start (it is empty when the Template opens with a token). Compared
    ' case-insensitively, consistently with the tokeniser - ParseTemplate matches "Prop[" with
    ' vbTextCompare, so a Template can legitimately hold a casing the canonical form does not.
    '
    ' The VISIBLE's own text goes into the new Template for every span that was EDITED (:698 keeps sSpan
    ' verbatim). A SURVIVING span does not: it is re-materialised CANONICALLY through TokenLiteral (:695),
    ' so this function can hand back a Template spelled "Prop[" for a visible that reads "prop[". That is
    ' not cosmetic - it is what let a lowercase token authored on an already-bound bearer sit unrendered
    ' for ever, until branch 2 was taught to recognise the never-rendered state case-insensitively.
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

        ' ANCHOR FIRST, literal search only as the fallback. Locating the closing literal blindly is
        ' wrong the moment a VALUE can contain it - values are copied verbatim, so "1 m 2" is a legal
        ' value under the Template "Prop[Len] m", and InStr would stop INSIDE the value. The span would
        ' then come back short, a purely static edit would read as a value edit, and the token would be
        ' released; with several tokens, one misaligned span misaligns every later one. When the known
        ' string is still sitting at the cursor, no guessing is needed at all.
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

        ' D7 - the span still CONTAINS the value intact, with the user's text beside it. This is the one
        ' place both edges of a span meet, and it generalises D6's leading half rather than sitting next to
        ' it: the value is untouched, so nothing has to be guessed about it; what has to be decided is
        ' whether the text beside it can be kept as static content while the token stays live.
        '
        ' WHY IT IS NOT ENOUGH TO LOOK AT THE ADDITION ALONE. The addition has TWO boundaries, not one -
        ' with the value, and with whatever already sat on its other side. Only the first is visible in the
        ' addition itself, so each test is handed the addition CONCATENATED WITH ITS CONTEXT (the literal
        ' before it, or the literal after it). That is not defensive padding, it is the fix for two real
        ' forges the addition alone cannot see:
        '   "T70Prop[x]m" + typing "x" before the value -> "T70x20m" reads 0x20, hexadecimal
        '   "T70Prop[x]m" + typing " " before the value -> "T70 20m" welds 70 and 20 into one number
        ' and it is equally what makes the ordinary case work: on "tProp[x]m", a space typed after the "t"
        ' is safe precisely BECAUSE the "t" is there to cut the reading - which only the context shows.
        '
        ' ONE side at a time, deliberately. The value is located by matching it against one END of the span,
        ' never by searching INSIDE it: a value of "1" inside a span of " 1 1 " has no defensible position.
        ' So an addition on both sides at once satisfies neither test and releases, exactly as today.
        '
        ' The ElseIf makes the PREFIX reading win when a span both starts and ends with the value (sEntry
        ' "1", sSpan "1 1"): the suffix reading is then never evaluated. Safety does NOT rest on that
        ' precedence - each branch is checked on its own merits, so whichever ran would have to pass - it is
        ' only what keeps the outcome deterministic instead of order-of-test dependent.
        '
        ' The right-hand test uses lits(i + 1), not sLit: sLit is the same literal in the USER's casing
        ' (matched with vbTextCompare). Casing turns no letter into a digit and no digit into a letter, so
        ' the two are interchangeable here. Left as is on purpose - "fixing" it would suggest a difference.
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

' True when sAnchor - the string the LAST rendering left in this span - is still sitting exactly at the
' cursor AND is immediately followed by the literal that closes the span. When that closing literal is the
' trailing EMPTY one, the anchor must instead run to the very end of the string: anything after it is text
' the user appended, which means the span is no longer just the anchor.
' The closing literal is always matched case-insensitively (the tokeniser's rule); the anchor's own
' comparison is the caller's call - binary for a VALUE, case-insensitive for a re-materialised token
' literal.
Private Function SpanAnchorsAt(ByVal sVisible As String, ByVal cursor As Long, ByVal sAnchor As String, ByVal sNextLit As String, ByVal lAnchorCompare As Long) As Boolean
    SpanAnchorsAt = False
    If Len(sAnchor) = 0 Then Exit Function
    If cursor < 1 Then Exit Function
    If cursor + Len(sAnchor) - 1 > Len(sVisible) Then Exit Function
    If StrComp(Mid(sVisible, cursor, Len(sAnchor)), sAnchor, lAnchorCompare) <> 0 Then Exit Function

    If Len(sNextLit) = 0 Then
        ' No closing literal, so nothing bounds the end of the span: the anchor normally has to run to the
        ' very end of the string. D6 relaxes that in ONE direction only - text APPENDED after the anchor is
        ' accepted when the boundary is provably safe (SuffixIsSafeAddition), which is what makes
        ' "135" -> "135m" keep its binding instead of releasing it. Everything else still requires
        ' end-of-string, so an in-place edit of the value ("13" -> "135") releases exactly as before.
        ' The caller needs no change: sSpan comes back as the anchor itself and AlignVisible's trailing
        ' "anything past the final literal is static text" line keeps the addition.
        If cursor + Len(sAnchor) = Len(sVisible) + 1 Then
            SpanAnchorsAt = True
        Else
            SpanAnchorsAt = SuffixIsSafeAddition(sAnchor, Mid(sVisible, cursor + Len(sAnchor)))
        End If
    Else
        SpanAnchorsAt = (StrComp(Mid(sVisible, cursor + Len(sAnchor), Len(sNextLit)), sNextLit, vbTextCompare) = 0)
    End If
End Function

' D6 - ASCII letter / ASCII digit, by CODE POINT. AscW, not Asc, so neither answer depends on the active
' code page: an accented letter, a Unicode digit or a fullwidth digit is none of these, and being none of
' these is always the REFUSING side of every test below.
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

' D6 - the whitelist test. True only for a character explicitly admitted: an ASCII letter, or one of the
' symbols in SAFE_BOUNDARY_SYMBOLS (see the admission criterion there). Anything else - digits, "." and ","
' obviously, but equally Chr(160), a narrow no-break space, a Unicode digit, a character nobody has thought
' of yet - is False, because it was never listed. That is the whole point: the refusal is the DEFAULT.
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

' D6 - may an addition AFTER the value be kept while the token keeps updating?
'
' Three conditions, all load-bearing:
'   1. the anchor is NUMERIC. A textual value is not covered, deliberately: "AB" -> "ABC" is as
'      indistinguishable as "13" -> "135", but there is no equivalent of IsNumericText to bound the damage,
'      so that domain keeps today's behaviour and gains no new risk.
'   2. the first character of the addition, ASCII spaces skipped, is on the whitelist. Spaces are skipped
'      rather than accepted because a space IS part of IsNumericText's alphabet - "135" + " 000" would
'      otherwise be kept and a later value of 20 would render "20 000", one plausible number. Skipping is
'      also why Chr(160) needs no mention: LTrim leaves it in place, and it is not on the whitelist.
'   3. the EXPONENT guard. A letter is admitted by 2, but "value + e5" reads as scientific notation, so a
'      whitelisted letter is refused when a digit follows it (an optional sign in between, which closes
'      "e-5" too). Written by SHAPE - digit, letter, digit - so it covers "0x20", "0b1", and any base
'      prefix nobody has named, instead of blacklisting e and x.
'
' What this does NOT do is guess what the user meant - that is impossible (appending "5" to "13" and
' retyping "135" produce the identical string). It bounds the CONSEQUENCE of being wrong: with the boundary
' provably unable to weld the addition onto a number, no future value can be silently forged. A wrong call
' shows "20m" instead of a frozen "135m" - visible, keeps the true value on screen, and undone by editing.
'
' sSuffix is EVERYTHING to the right of the value, not just what the user typed: D7's caller appends the
' closing literal, so an addition of " " on the Template "tProp[x]5m" is judged as " 5m" and refused - the
' welding it would allow lives one character past what the user typed. Where nothing follows (a terminal
' token, SpanAnchorsAt's caller) there is no context to append and the addition IS the whole right side.
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

' D6, mirror image - may an addition BEFORE the value be kept? Same three conditions, opposite end: the LAST
' character of the addition, trailing ASCII spaces skipped, must be on the whitelist, and a letter there is
' refused when a digit PRECEDES it ("1e" + 20 -> "1e20", "0x" + 20 -> "0x20"). The space skip matters as
' much here: a prefix of "1 " ends in a space, and keeping it would render "1 20" once the value moves.
'
' sPrefix is EVERYTHING to the left of the value, not just what the user typed - D7's caller prepends the
' opening literal. That is what decides the two cases the typed text alone cannot tell apart: on
' "tProp[x]m" a typed space becomes "t ", whose "t" cuts any numeric reading, so it is KEPT; on
' "T70Prop[x]m" the same space becomes "T70 ", ending on a digit, so it is REFUSED - and the "x" that would
' spell "T70x20" is caught by the same concatenation through the digit-letter-digit guard.
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

' D8 - the RAW frontier of a literal: its first / last two characters, BEFORE any trimming. It answers one
' question only - did the text touching the value change? - and the rawness is the whole point. RTrim would
' hide the very edit that matters: an authored "T70" with a space typed after it trims back to "T70" and
' would read as untouched, while the frontier truly went from "70" to "0 ".
Private Function RawHead(ByVal s As String) As String
    RawHead = Left(s, 2)
End Function

Private Function RawTail(ByVal s As String) As String
    RawTail = Right(s, 2)
End Function

' D8 - could the END of this literal be the beginning of a number? Answered on the trailing run of
' number-capable characters taken WHOLE - bounded by the data, never by a constant.
'
' This exists because D6's guards look a FIXED distance back (a character, then one more for the exponent
' test), which is exactly as far as an ADJACENT addition could ever reach. D8 accepts an edit anywhere, so
' "0x" typed at the far start of a literal ending in "A1" spells "0xA1", and a future 20 renders "0xA120" -
' hexadecimal, four characters from the frontier the guards inspect. Any constant depth is the blacklist
' mistake wearing new clothes: an enumeration over an open space.
'
' The closed form, and the reason no base prefix has to be named: A NUMBER STARTS WITH A DIGIT. Decimal,
' 0x, 0b, 0o - all of them. So a trailing run is number-capable exactly when its FIRST character is a
' digit, and every base notation is covered by that one sentence instead of by a list of the ones we
' happen to know.
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

' D8 - is the text now sitting to the LEFT of a value acceptable? Three questions, and the ORDER is the
' design:
'   1. did the edit CREATE the possibility of a base literal where there was none? That is the only
'      non-local forge, and it is judged first because the frontier test below cannot see it: typing "0x"
'      in front of an authored "A1" leaves the last two characters untouched.
'   2. did the frontier change at all? An untouched literal is not judged - it is the user's authored text,
'      and D7's own claim is that the rule may not be stricter than what the Template already spells. That
'      is what lets "b" typed at the very start of "Zone 70Prop[Len]m" keep its binding: "70" still touches
'      the value, nothing new does.
'   3. only then, D6's guard on the whole current literal.
' The frontier is the RAW last two characters, so an edit that only LOOKS harmless after trimming - "T70"
' becoming "T70 " - still counts as a change and is judged.
Private Function LeftContextIsSafe(ByVal sAuthored As String, ByVal sCurrent As String, ByVal sValue As String) As Boolean
    LeftContextIsSafe = True

    ' The user deleted the literal outright, so the value now starts the text. EMPTY is the one context
    ' that provably cannot weld: there is no character to weld WITH. Note this is the exact opposite of what
    ' empty means to the shared predicates, where it says "no addition was made, this branch should not have
    ' run" - which is why they refuse it, and why this case is handled HERE and never in them.
    If Len(sCurrent) = 0 Then Exit Function

    LeftContextIsSafe = False
    If NumericTailIsPossible(sCurrent) And Not NumericTailIsPossible(sAuthored) Then Exit Function

    LeftContextIsSafe = True
    If RawTail(sAuthored) = RawTail(sCurrent) Then Exit Function

    LeftContextIsSafe = PrefixIsSafeAddition(sCurrent, sValue)
End Function

' D8 - mirror image, for the text now sitting to the RIGHT of a value.
' No base-literal guard on this side, and that is not an omission: a base literal needs its PREFIX ("0x")
' in front of the digits, so the family can only ever form to the LEFT of a value. What can form on the
' right is an exponent, and SuffixIsSafeAddition already closes it.
Private Function RightContextIsSafe(ByVal sAuthored As String, ByVal sCurrent As String, ByVal sValue As String) As Boolean
    RightContextIsSafe = True

    ' Same as the left half: an emptied literal leaves the value at the end of the text, with nothing to
    ' weld onto. A gap emptied BETWEEN two values would make two tokens adjacent instead - that one is not
    ' waved through here, it is caught by the well-formedness half of ReauthoredTemplateIsSound, which is
    ' the single place this mechanism checks what it re-authored.
    If Len(sCurrent) = 0 Then Exit Function
    If RawHead(sAuthored) = RawHead(sCurrent) Then Exit Function

    RightContextIsSafe = SuffixIsSafeAddition(sValue, sCurrent)
End Function

' D8 - the ONE check that makes a re-authored Template safe to store, and it replaces three separate
' arguments with a single verification.
'
' The mechanism below promotes ALL non-value text to literal, so a user who happens to type "Prop[Other]"
' into a bound text gets it promoted too. That is the round-8/10 WEDGE: the entry would carry a token set
' its LastValues does not match, EntryIsConsistent would refuse it as vandalised metadata on every pass
' afterwards, and nothing recovers it.
'
' Rather than argue the three properties separately, verify them:
'   1. the re-authored Template is well-formed (no duplicate, no adjacent tokens - both newly reachable
'      here, since ordinary user text becomes literal);
'   2. it carries EXACTLY the tokens we put in it - a "Prop[...]" that came from the user's own typing
'      changes the count and is caught;
'   3. THE FIXED POINT: expanding it with the very same LastValues reproduces the visible text byte for
'      byte. This is the property the whole engine rests on, and it is checked, not inherited.
' Any failure returns False and the caller falls back conservatively, which is today's behaviour.
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

' D8 - alignment by VALUE RECOGNITION, the fallback for everything the literal walk cannot follow.
'
' AlignVisible locates a value by IMPOSING positions: lits(0) must sit at offset 1, the value must sit at
' the cursor, the closing literal must sit right behind it. Every geometry that broke between rounds 35 and
' 41 was a position that moved - text typed before the leading literal, after the trailing one, between a
' literal and the value. This function inverts the search: find each LastValues entry in the visible, and
' whatever lies between the matches BECOMES the new literal. Positions stop mattering, so static text is
' free to change anywhere, in any quantity, on any number of tokens.
'
' Three conditions, all refusals rather than guesses:
'   - every token must have a NON-EMPTY value. The unset state renders as the token's own literal text and
'     is left to AlignVisible, which already handles it.
'   - each value must appear EXACTLY ONCE. Two occurrences have no defensible choice between them, and
'     picking one would move the token - a silent corruption far worse than releasing. A short value that
'     also occurs in the static text ("3" in "3x240") therefore releases; that is a known limitation, not
'     an oversight, and it is where AlignVisible's literal anchoring still earns its keep.
'   - both frontiers of every value must be safe (see LeftContextIsSafe / RightContextIsSafe).
'
' Runs ONLY where AlignVisible has already failed or survived nothing, so no path that works today changes.
'
' THE NON-LOCAL FORGE, and why it is closed rather than accepted. Widening the surface from "adjacent
' addition" to "edit anywhere" reopens one family the D6/D7 guards cannot see: a base literal. Typing "0x"
' at the far start of a literal ending in "A1" spells "0xA1", and a future 20 renders "0xA120" - four
' characters from the frontier those guards inspect, and with the last two unchanged, so neither the guard
' nor the frontier test fires. It was briefly accepted as residual; NumericTailIsPossible closes it
' instead, on ONE sentence that needs no base prefix enumerated - a number starts with a digit.
' The rule is differential, like D7's claim: a base literal that was ALREADY possible in the authored
' Template is not held against the user (that is the arbitrated "authored, not forged" class), only one the
' edit newly makes possible. The case that justifies the differential is not some exotic one - it is
' "Zone 70Prop[Len]m", whose trailing run ALREADY starts with a digit. Judged absolutely, it would be
' refused with no edit anywhere near the value: the mandate's own headline case.
'
' COMPLETENESS, by argument rather than by the test corpus: a value is IsNumericText - digits, space, comma,
' point - so it can never supply the "x". A base reading therefore needs the "0x" inside the LITERAL, and
' every character able to continue that reading (x, hex digits, digits) is number-capable, so the run
' holding that "0" necessarily reaches the end of the literal and its first character IS that "0". Any
' non-capable character in between breaks the run and the reading together ("0xA1 " + 20 reads "0xA1 20").
' There is no gap between the two.
' Public for the same reason ExpandTemplate and AlignVisible are: it is a read-only test seam (AC12), so
' the decision table below can be asserted without a DGNLib. It writes nothing and touches no element.
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

' One drain target: skip it when this drain already ran it, re-fetch it, then run the state machine.
'
' The done-list is what stops an element being processed twice in one drain. It legitimately turns up
' twice - once as a noted source, once as another source's group sibling - and the second run is not
' merely wasted work: it reads the element through a handle captured BEFORE the first run rewrote it, sees
' the pre-write text, matches neither expansion, and releases the binding as if the user had retyped it.
' That is the "one of the two texts loses its binding" failure, and it needs a group with TWO bound texts
' plus a value that keeps moving - which is why one text, however fast it is edited, never shows it.
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
' this module: never read text through a handle you did not just fetch.
'
' Every element handle reaching this engine was captured earlier by someone else - the idle batch
' materialises all of its elements before processing the first one, moDirty stores them at value-write
' time, Link.GetLink scans before the walk writes anything, the bind key-in loops over a selection set
' taken up front - and a handle keeps serving the TEXT it was captured with. The ITEM side does not
' behave that way: CustomPropertyHandler refreshes it from the file on every access. Stale text plus
' fresh values is the one combination the state machine cannot survive, because it matches neither
' expansion and branch 3 is entitled to call that a user edit.
'
' The same staleness is why StringsInEl.UpdateTextLines re-fetches after every sub-write and why
' ElementInProcesse stores ids rather than elements. Best effort: an element that cannot be re-fetched
' keeps the handle it had, which is no worse than before.
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

' Branches 1-3, per stored entry. texts() holds the whole text of every bearer, indexed by SubId, from
' the single walk the caller already paid for.
'
' Write ORDER is load-bearing and enforced here: every TEXT write happens first, the ONE metadata write
' last. Metadata-first with a failed text write would leave the visible and LastValues disagreeing, so
' the next pass would match neither expansion and release the binding spuriously. If the metadata write
' fails after the text landed, every text written in this pass is restored (best effort) and NOTHING
' transitions - the in-memory entry changes are simply discarded.
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
        ' Branch 3 transitioned an entry WITHOUT writing text, so the bearer may still show a literal
        ' token. Note it only HERE, after the metadata write succeeded: the drain re-reads the entry from
        ' the file, so noting it before the write would re-render it against the OLD Template. Once per
        ' element and per pass, never once per entry - moDirty would otherwise carry the same element as
        ' many times as it has re-authored sub-texts.
        ' Terminates: the drain runs the FULL state machine, branch 2 renders and WRITES, and a written
        ' entry never sets bRepaint. DrainRepaintHop is re-entrance-guarded and clears its list, so a note
        ' added from inside a drain is dropped rather than replayed.
        NoteDirtyGroup oEl
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.RenderBoundElement"
End Sub

' The three branches for ONE entry. Returns ENTRY_UNCHANGED / ENTRY_UPDATED / ENTRY_DROP.
' SubId is the ordinal ResolveAllSubIds settled for this entry (-1 = refuse); it may differ from the
' stored one, which is what a drift relocation means.
' wSubId + wPrev report a text write that DID land, so the caller can undo it if the metadata write
' then fails (wSubId stays -1 when nothing was written).
' oEl is ByRef all the way down to StringsInEl.SetTextAtSubId, so the handle refreshed after a
' sub-element Rewrite reaches the caller instead of dying in a local copy.
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
    ' The relocation is held in a LOCAL and committed only once the entry has proved usable. Committing
    ' it here would return ENTRY_UPDATED out of every "skip, never transition" exit below - the empty
    ' read, the vandalised metadata, the illegal value - and so write metadata for an entry the state
    ' machine has just declared unusable.
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

    ' ExpandTemplate fails OPEN (it returns the Template unchanged), and it reaches COM through
    ' ParseTemplate -> IsKnownProperty. Comparing the visible against a garbage expansion would miss
    ' branches 1 and 2 and land in branch 3, the only destructive one - a transient library-read fault
    ' would then become a permanent binding release. Skip the entry instead, exactly as an empty or
    ' faulted read does.
    sExpCur = ExpandTemplate(ents(idx).Template, curNames, curValues, nCur, True, bOkCur)
    If Not bOkCur Then Exit Function
    sExpLast = ExpandTemplate(ents(idx).Template, ents(idx).ValNames, ents(idx).ValValues, ents(idx).nVals, True, bOkLast)
    If Not bOkLast Then Exit Function

    ' The entry is READABLE, CONSISTENT and its values are legal: only now may a relocation be persisted
    ' (edge #17 - "on success, update the stored SubId"). An entry that did NOT drift still returns
    ' ENTRY_UNCHANGED here, so branch 1 stays the strict no-op AC2 requires.
    If bRelocated Then
        ents(idx).SubId = SubId
        RenderEntryOnElement = ENTRY_UPDATED
    End If

    ' --- BRANCH 1: up to date. STRICT no-op - no text write, no Rewrite, no metadata write. The loop
    ' terminator, and the reason a re-queued unchanged element costs nothing.
    If sVisible = sExpCur Then Exit Function

    ' --- BRANCH 2: the visible still matches the LAST rendering, so the VALUES moved. Re-render.
    '
    ' Second test: an entry authored but NEVER rendered has no last rendering at all - its visible IS its
    ' Template, in whatever casing the USER typed, while ExpandTemplate always re-materialises the
    ' canonical "Prop[" prefix. A "prop[Name]" typed into an ALREADY-BOUND bearer is authored verbatim by
    ' the top-up scan with an empty LastValues entry, so both expansions come back capitalised, neither
    ' matches binary, and the entry lands in branch 3 - which rewrites the Template to the canonical form,
    ' resets the value to empty and writes NO text. Every later pass repeats it identically: the token
    ' never renders, and nothing is reported. The first-author path never sees this because it expands and
    ' writes straight from the current values, which is why the SAME token renders on an unbound bearer.
    ' Compared case-insensitively, like every other INTERPRETATION in this module (AlignVisible's token
    ' literals); IDENTIFICATION - ResolveAllSubIds - stays binary. This also recovers an entry already
    ' trapped by an earlier pass, whose stored Template is canonical while its visible is not.
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
    ' Per-token release. No DATA property is touched, and no text is written: the user's own text IS
    ' the new Template.
    If AlignVisible(sVisible, ents(idx).Template, ents(idx).ValNames, ents(idx).ValValues, ents(idx).nVals, sNewTemplate, newNames, newValues, nNew) Then
        ' The re-authored Template is re-validated exactly like a first author: the state machine may
        ' never produce a Template that breaks a v1 invariant.
        If TemplateIsWellFormed(sNewTemplate, True) Then
            ' TemplateIsWellFormed is True for a token-FREE string too - nothing in it can be malformed -
            ' so it cannot be the only gate. Edge #5 (the user retypes the whole text) releases the last
            ' token and lands here with ZERO tokens left, and storing THAT as a live entry would be a
            ' one-way trap: ARES_Render would stay attached for ever, CountLiveEntries would never reach
            ' 0 so the detach would never fire, IsRenderBound would stay True, and AutoLengths' Branch 1
            ' would skip a text that no longer carries a single token - with no unbind key-in in Phase 1
            ' to recover it. Release the entry instead, the same nTok = 0 test ApplyConservativeFallback
            ' already makes.
            If ParseTemplate(sNewTemplate, newLits, newToks, nNewTok, True) Then
                If nNewTok > 0 Then
                    ' A token the user TYPED into an already-bound text is in the new Template but has no
                    ' LastValues entry, and EntryIsConsistent demands the two sets match EXACTLY - the
                    ' entry would be read as vandalised metadata from then on and never render again, with
                    ' no way back. Give every token the new Template carries an entry, empty for the ones
                    ' that just appeared: that is what ApplyConservativeFallback already does, and an empty
                    ' entry renders as the literal token until a value shows up and branch 2 takes over.
                    ' Nothing is resurrected here - a token the user WIPED is no longer in the visible, so
                    ' it is not in newToks either.
                    For k = 0 To nNewTok - 1
                        If Not HasValueEntry(newToks(k), newNames, nNew) Then
                            AppendValue newToks(k), "", newNames, newValues, nNew
                        End If
                    Next k
                    ents(idx).Template = sNewTemplate
                    SetEntryValues ents, idx, newNames, newValues, nNew
                    ' This branch writes NO text on purpose - the user's own text IS the new Template - so
                    ' it leaves the bearer showing whatever was typed, which for a re-typed token is the
                    ' LITERAL rather than the value. The state machine's own answer is "the very next pass
                    ' renders it (branch 2)", and that is true; what is not true is that a next pass ever
                    ' comes: the text stayed unrendered until the user edited it a SECOND time. The chain
                    ' most likely responsible - no text write, so no Rewrite, so no change event, so no
                    ' re-queue - is OBSERVED, NOT PROVEN (story 15-2, Task 0(b)): bulk detection,
                    ' ShouldQueueElement or idle-write timing would look identical from here. The fix does
                    ' not rest on that diagnosis, which is the point of doing it this way.
                    ' Ask for the repaint hop instead: the caller notes the element once the metadata is
                    ' safely written, and the drain that ProcessElement already runs re-enters the full
                    ' state machine on it, where branch 2 renders it in this same pass.
                    bRepaint = True
                    RenderEntryOnElement = ENTRY_UPDATED
                Else
                    ' The alignment SUCCEEDED and concluded that nothing survived - a deliberate release,
                    ' not a failure to understand the text. Until now it was the only way a binding could
                    ' die without the user being told anything at all, which is why it gets its own status
                    ' rather than borrowing ReportAmbiguous below: that one says "the text could not be
                    ' matched", which is the opposite of what happened here, and both being one-shot the
                    ' less accurate message would win whichever fired first.
                    ReportBindingReleased
                    RenderEntryOnElement = ENTRY_DROP
                End If
                Exit Function
            End If
        End If
    End If

    ' --- BRANCH 3b (D8): the literal walk could not follow the text, so stop following literals.
    ' AlignVisible fails whenever a position moved - text typed before the leading literal, or a static word
    ' edited anywhere - even though the rendered VALUE is still sitting there untouched. Try to recognise
    ' the values themselves; everything else in the text then simply becomes the new static content.
    ' This runs ONLY here, after AlignVisible has already declined, so no path that works today is affected.
    If AlignByValues(sVisible, ents(idx).Template, ents(idx).ValNames, ents(idx).ValValues, ents(idx).nVals, sNewTemplate, newNames, newValues, nNew) Then
        ents(idx).Template = sNewTemplate
        SetEntryValues ents, idx, newNames, newValues, nNew
        ' Same repaint hop as branch 3: no text is written here either - the user's own text IS the new
        ' Template - so the caller must re-enter the state machine for branch 2 to render this pass.
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

' Conservative outcome: the visible text BECOMES the Template, so the "Prop[Name]" substrings the user
' left verbatim stay live (with an EMPTY LastValues entry - they are unset until the next value read)
' and everything else is static. That state converges: the very next pass either renders the values
' (branch 2) or finds itself already correct (branch 1).
'
' When the visible is not a well-formed Template (the user typed the same token twice, or two adjacent
' ones), NO valid binding can be stored for it - and storing a token-free Template is not an option
' either, since a Template is re-parsed from its string and any "Prop[Known]" in it is always a token.
' The entry is therefore RELEASED outright, which is also what keeps this branch from churning the
' metadata on every single pass.
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
' not the element: a cell header carries ONE ARES_Render item holding an entry per bound sub-text, so
' "this element is already bound" says nothing about whether a given sub-text is. Both the first-author
' scan and the top-up scan on an already-bound bearer run this one function, so the rules cannot drift.
'
' The hybrid policy (§6.2) in full: at least one valid token, a well-formed Template (no duplicate or
' adjacent tokens), a bearer allowed to carry tokens, and EVERY token naming a property ALREADY ATTACHED
' to the element - attachment being the intent signal that story 15-1's convergent "@" pull makes reliable.
' bKeepSource / nFree carry the source-preservation rule (see FeedsCellSource). They live HERE, at the one
' place an entry is ever created, and not in the callers: the round-8 regression was let through by a
' guard that sat in only ONE of this function's two callers, so the two authoring paths diverged again on
' exactly the rule that mattered. nFree is the number of sub-texts still outside the exclusion set, and it
' is decremented here on every successful author, so both callers stay honest without repeating anything.
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

    ' SOURCE PRESERVATION, applied after the token test on purpose: a bearer carrying no token was never a
    ' candidate, is exactly what we want left behind as the source, and must not draw a warning. Only a
    ' bearer that WOULD have been authored can take the last free sub-text away.
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

' Does this element FEED a Cell* calc source? If it does, it must keep at least ONE sub-text out of the
' exclusion set, and TryAuthorBearer enforces that.
'
' GetExcludedSubIds hides every rendered sub-text from GetConcatenatedText, so binding the LAST unbound one
' leaves the CellText source with nothing to read: it resolves to "", calc empties the very properties the
' cell displays, and the renderer then shows them all as literal tokens - the containment destroying the
' data it exists to protect. There is no repair on the READ side either: once every sub-text is rendered,
' any non-empty read necessarily contains rendered text, so every possible fallback IS the ratchet. The
' exclusion is structurally incompatible with "all sub-texts bound", which makes refusing the last one the
' only lever available - the same conservative call §4.4a already makes for a TextNode in a matched cell,
' through the same IsTriggerCell superset.
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
' the same authoring rule. Returns how many entries were added.
'
' Without this, first-authoring is gated on the ELEMENT being unbound while the unit of binding is the
' SUB-TEXT, and the two disagree the moment a cell holds more than one token. The first sub-text to
' qualify attaches ARES_Render to the cell header; from then on every pass takes the bound branch, which
' only ever loops the entries already stored - so a second sub-text whose property arrives later is never
' looked at again, by any path, including the bind key-in. It stays inert for ever, exactly as if its
' token had never been typed.
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

    ' Entries just read from metadata are never Dropped, so the test is defensive - but it states the
    ' rule: only a LIVE entry owns its sub-text. The case this really serves is the reported one - a
    ' sub-text whose token names a property that was not attached YET at first author. On an unbound
    ' element the author scan simply retries every pass; on a bound one, nothing retried it at all.
    '
    ' An UNRESOLVED live entry (subIds = -1) is the dangerous case: its sub-text is unknown, so leaving it
    ' unclaimed would let this scan author a DUPLICATE for the very text that entry drives - and the
    ' re-resolution that follows hands the ordinal to the newcomer, stranding the original in permanent
    ' drift as a dead record in the blob. The picture is incomplete, so author nothing this pass; the
    ' entry already reports drift, and once it resolves the top-up resumes. Refuse rather than guess.
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

' Branch 4 - the first author, hybrid policy. bManual = the BindPropertyRender key-in (which reports why
' it refused); the automatic path stays quieter but uses the exact same rules.
'
' A bearer is bound only when its visible text carries at least one valid token AND every one of those
' tokens names a property ALREADY ATTACHED to the element. Attachment is the intent signal - it is what
' story 15-1's convergent "@" pull makes reliable - so nothing is ever bound by accident.
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

    ' One decision per BEARER, taken by the shared authoring rule - the same one the top-up scan runs on an
    ' already-bound element, so the two can never drift apart. That includes SOURCE PRESERVATION: a first
    ' author binds every qualifying sub-text in ONE pass, so a fresh cell whose sub-texts all carry valid
    ' tokens would otherwise empty its own CellText source on the spot - the round-8 failure through the
    ' other door. Nothing is bound yet here, so every bearer is still free.
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
' it can reach the file. It does NOT refuse an expansion shaped like an ACTIVE legacy trigger - writing
' "(12.3m)" is the flagship case, and AutoLengths stands down via Branch 1's IsRenderBound skip instead.
' Returns True only when the sub-text now reads sNew.
'
' oEl is ByRef, and the WHOLE chain above it (RenderBoundElement / TryFirstAuthor / RenderEntryOnElement)
' is too, on purpose: StringsInEl.SetTextAtSubId re-fetches the element after the sub-element Rewrite that
' makes the handle stale, and VBA's ByVal on an object copies the REFERENCE - the refreshed handle would
' die in this procedure's local copy. It has to reach the caller, because on a cell with two rendered
' sub-texts the SECOND write and the closing WriteRenderMetadata both run off that same handle.
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
' Values are copied VERBATIM: a non-string typed value (a date or boolean from a hand-authored lib) is
' NOT converted - it yields the literal token plus a status, because a locale-dependent conversion is
' exactly what would make two stations rewrite each other's text forever.
' Returns False only when a value carries a reserved delimiter or a line break - the whole entry is then
' skipped, so no VALUE can ever make the metadata unparseable.
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

' Locate the sub-text EVERY entry drives, in TWO PASSES over the whole entry list. subIds(i) comes back
' as the resolved ordinal, or -1 when the entry must be refused rather than written blind.
'
' The two passes are what make the result independent of entry ORDER, and that is the whole point:
'   Pass A - every entry that can PROVE nothing drifted (the text at its stored ordinal still IS its
'            computed last rendering, or its Template) takes that ordinal and claims it.
'   Pass B - only then do the still-unresolved entries scan the rest of the cell for their text, over
'            the ordinals nobody claimed, and fall back to their own stored ordinal when in range.
'
' A single pass gets this wrong in both directions. Resolving entries one at a time, an entry whose
' sub-text the user just EDITED no longer matches at its own ordinal, so its global scan finds a SIBLING
' whose rendering happens to read the same string and steals it - the sibling is then refused, and the
' user's edit is interpreted against the wrong text instead of reaching branch 3. Preferring the stored
' ordinal outright would fix that and break edge #17 instead: after a cell redefinition inserts a
' sub-element, every stored ordinal is still IN RANGE while designating someone else's text, and the
' relocation would never fire. Pass A settles all the certain cases first, so pass B only ever scans what
' is genuinely unaccounted for - and two entries that both drifted still both relocate correctly.
'
' Comparisons are BINARY on purpose. This is IDENTIFICATION, not interpretation: two sub-texts differing
' only in case are two different texts, and a value's casing is data (values are copied verbatim). Both
' sides of the "last rendering" test are produced by ExpandTemplate, so the canonical "Prop[" prefix
' cannot make them disagree; AlignVisible, which interprets rather than identifies, is the one that
' matches token literals case-insensitively.
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

' Pass B, for the entries pass A could not settle:
'   1. a sub-text elsewhere carrying this entry's last rendering, then one carrying its Template - the
'      cell was redefined and the ordinals shifted, so the entry follows its text (edge #17);
'   2. otherwise its own stored ordinal if that is still in range and unclaimed - the ordinal is fine and
'      the text simply no longer matches because the USER edited it, which is what branch 3 interprets;
'   3. otherwise -1: refuse rather than write blind into a sub-text that is not ours.
' Every step skips ordinals another entry already owns.
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

' Defensive bearer guard: the bearer is always a TOP-LEVEL model element, never a cell component. The
' Ouroboros exclusion is keyed on the CELL that carries ARES_Render, so a sub-text bound DIRECTLY would
' compute no exclusion at all and the rendered span would feed the very CellText value that governs it.
'
' In normal use this state is UNREACHABLE - attaching a custom property to a cell component without
' dropping the cell first is not something the UI offers - but BindPropertyRender is a new key-in whose
' selection semantics carry none of the native attach command's guarantees, so the invariant is enforced
' rather than assumed.
'
' It refuses SILENTLY, and that is deliberate: none of the 20 sanctioned i18n keys describes this case
' truthfully (reusing RenderTextNodeInCellRefused would tell the user something about multi-line texts in
' cells that is simply not what happened), and inventing a key is barred until the spec rules on it. A
' wrong message is worse than none for a state the user cannot reach; the day it becomes reachable, it
' needs its own key.
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

' v1 refuses to author a token inside a TEXTNODE that belongs to a cell fed by an active group source:
' the exclusion granularity is one whole bearer, and a TextNode IS one bearer, so a single token line
' could not be excluded without also hiding the cell's other lines from the calc value. A plain
' sub-TextElement is excludable and therefore allowed.
'
' The "fed by a group source" test reuses PropertyCalculation.IsTriggerCell, which is broader than
' CellText alone (it covers every pushable Cell* source). Deliberately conservative: over-refusing a bind
' costs a status message, under-refusing corrupts a value.
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

' Read and parse the ARES_Render metadata of an element. False = attached but unusable (unreadable,
' unknown schema, or malformed) - the caller must refuse, never guess. True with nEnts = 0 is the legal
' "freshly attached, nothing stored yet" state.
'
' EVERY False path reports its OWN status here, and callers must not add one: a schema mismatch has a
' dedicated key (edge #24) and an unconditional generic "unreadable" from the caller would overwrite it
' on the status bar, defeating the point of having it.
'
' EVERY access names ItemName AND LibraryName explicitly, and suppresses the single-property fallback.
' Two independent traps sit here: CustomPropertyHandler defaults an omitted ItemName to the PROPERTY
' name (the "ItemType name IS the property name" convention that ARES_Render deliberately breaks), which
' would make every read return Null silently; and its fallback returns "the first property that yields a
' value", which on this 2-property ItemType hands back Entries when SchemaVersion is empty.
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

' A Template is well formed when no property is tokenised twice (v1 rule: ONE token per property per
' text) and no two tokens are adjacent. An interior EMPTY literal is what adjacency produces, and it
' leaves the span boundary undefined - which is precisely what the per-token release needs in order to
' decide what the user edited.
' Public (read-only) as a test seam, like ExpandTemplate and AlignVisible: the duplicate/adjacent
' refusals are acceptance criteria and VBA cannot call a Private proc from UnitTesting.
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

' Third self-disable condition, and the one that is not a refusal but a FAULT LOOP. Both writers restore
' the text they wrote and return, leaving the persisted state byte-identical to what they found - so the
' next pass takes exactly the same path, emits exactly the same Rewrites, and each of those re-queues the
' element. Branch 1, the declared loop terminator, is never reached because nothing is ever retained.
' Log ONE English line and go inert for the session, the same shape as the schema/library conditions. The
' user's recovery is to fix the DGNLib and reopen; RefreshRenderCaches clears it for the tests.
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
