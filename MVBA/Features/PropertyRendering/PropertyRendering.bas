' Module: PropertyRendering
' Description: The SOLE text writer of ARES. Displays a custom property's value inside a text by fully
'              replacing a "Prop[Name]" token; the binding lives in hidden ARES_SYS metadata, never in
'              the visible text or the graphic group. Full mechanism (binding storage, the 4-branch
'              release state machine, value semantics): see _bmad/docs/property-rendering-mechanics.md.
'              This module is the Core / public API / orchestration surface. The wider feature is split
'              across PropertyRendering_Types (RenderEntry + array helpers), PropertyRendering_TemplateModel
'              (ExpandTemplate/AlignVisible/AlignByValues/TemplateIsWellFormed, ParseTemplate, and the D6-D8
'              forge-protection guards), PropertyRendering_StateMachine (render branches 1-3b),
'              PropertyRendering_Authoring (bind/first-author/top-up), PropertyRendering_Serialization
'              (ARES_Render metadata read/write) and PropertyRendering_Reporting (one-shot status).
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (global ARESConfig), CustomPropertyHandler, PropertyTagging,
'               StringsInEl, Link, LangManager, ErrorHandlerClass (global ErrorHandler),
'               CallStackClass (global CallStack), PropertyRendering_Types, PropertyRendering_TemplateModel,
'               PropertyRendering_StateMachine, PropertyRendering_Authoring, PropertyRendering_Serialization,
'               PropertyRendering_Reporting
'
' Never writes a USER property value or attaches/detaches anything itself (delegated to PropertyTagging).
' Any SetPropertyValueToElement on the "ARES" library inside this module is a review BLOCKER.

Option Explicit

' The two self-disable conditions also write an English log line. Those flags are SESSION-scoped, not
' per-element. Third SESSION-scoped self-disable: guards against an unbounded write/restore fault loop,
' not a normal refusal. Full rationale: see "Session-scoped self-disable guards" in
' property-rendering-mechanics.md.
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

    PropertyRendering_Reporting.ResetOneShots

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
        PropertyRendering_Reporting.ReportLocked
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
        PropertyRendering_Reporting.ReportLibraryMissing
        GoTo DrainAndExit
    End If

    If CustomPropertyHandler.IsItemAttachedToElement(oEl, ARES_ITEM_RENDER, ARES_NAME_LIBRARY_SYS) Then
        PropertyRendering_StateMachine.RenderBoundElement oEl, texts, nBearers
    Else
        PropertyRendering_Authoring.TryFirstAuthor oEl, texts, nBearers, False
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
    If Not PropertyRendering_Serialization.ReadRenderMetadata(oCell, ents, nEnts) Then Exit Function
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
    PropertyRendering_Reporting.ResetOneShots

    If El Is Nothing Then Exit Function

    ' Self-disabled after a failed metadata write - a refusal the user asked for, so it must stay visible.
    ' See "First-author key-in visibility (BindElement)" in property-rendering-mechanics.md.
    If mbWriteDisabled Then
        PropertyRendering_Reporting.ReportMetadataUnreadable
        Exit Function
    End If

    ' Same reason as ProcessElement: the key-in loops over a selection set materialised before the first
    ' bind, so element k's render can leave element k+n's handle serving a pre-write text.
    Set El = FreshHandle(El)
    If El Is Nothing Then Exit Function

    If Not IsRenderableType(El) Then Exit Function

    If El.IsLocked Then
        PropertyRendering_Reporting.ReportLocked
        Exit Function
    End If

    If Not IsSysLibraryPresent() Then
        PropertyRendering_Reporting.ReportLibraryMissing
        Exit Function
    End If

    Dim ids() As Long
    Dim texts() As String
    Dim nBearers As Long
    nBearers = StringsInEl.EnumerateTextSubIds(El, ids, texts)
    If nBearers <= 0 Then Exit Function

    ' Already bound: nothing to author, just bring it up to date through the full state machine.
    If CustomPropertyHandler.IsItemAttachedToElement(El, ARES_ITEM_RENDER, ARES_NAME_LIBRARY_SYS) Then
        PropertyRendering_StateMachine.RenderBoundElement El, texts, nBearers
        BindElement = True
        Exit Function
    End If

    BindElement = PropertyRendering_Authoring.TryFirstAuthor(El, texts, nBearers, True)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.BindElement"
    BindElement = False
End Function

'######################################################################################################################
'                                          STATE MACHINE ENTRY / DRAIN
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
        PropertyRendering_StateMachine.RenderBoundElement oEl, texts, nBearers
    Else
        PropertyRendering_Authoring.TryFirstAuthor oEl, texts, nBearers, False
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering.RenderOneElement"
End Sub

'######################################################################################################################
'                                   METADATA - SUBID RESOLUTION
'######################################################################################################################

' Locate the sub-text EVERY entry drives, in TWO PASSES over the whole entry list (Pass A: certain matches
' only; Pass B: relocation scan). Independent of entry ORDER, and comparisons are BINARY (identification,
' not interpretation). Full rationale: see "SubId resolution - the two-pass algorithm" in
' property-rendering-mechanics.md.
' Public: called from Module D (StateMachine)'s RenderBoundElement, not just internally.
Public Sub ResolveAllSubIds(ByRef ents() As RenderEntry, ByVal nEnts As Long, ByRef texts() As String, ByVal nBearers As Long, ByRef subIds() As Long)
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

    If texts(SubId) = PropertyRendering_TemplateModel.ExpandTemplate(ents(idx).Template, ents(idx).ValNames, ents(idx).ValValues, ents(idx).nVals) Then
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
    sExpLast = PropertyRendering_TemplateModel.ExpandTemplate(sTemplate, ents(idx).ValNames, ents(idx).ValValues, ents(idx).nVals)

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

' Safe "array has at least one element" check (mirrors PropertyCalculation.HasElements). UBound returns
' -1 for an empty array and RAISES for an uninitialised one.
Private Function HasElements(ByRef arr() As element) As Boolean
    On Error Resume Next
    HasElements = False
    If UBound(arr) <> -1 Then HasElements = True
    On Error GoTo 0
End Function

' Third self-disable condition, and the one that is not a refusal but a FAULT LOOP guard. Full rationale:
' see "Session-scoped self-disable guards" in property-rendering-mechanics.md.
' Public: called from Module D (StateMachine) and Module E (Authoring) after a metadata write failure.
Public Sub DisableAfterWriteFailure()
    On Error Resume Next
    If mbWriteDisabled Then Exit Sub
    mbWriteDisabled = True
    ErrorHandler.HandleError "Property rendering: ARES_Render metadata write failed, rendering disabled for this session", 0, "", "PropertyRendering.WriteRenderMetadata"
End Sub

' Membership test against the cached DGNLib property names. Matching is case-insensitive, aligned with
' the rest of the system. Public: called from Module C (TemplateModel)'s IsAcceptableTokenName.
Public Function IsKnownProperty(ByVal sName As String) As Boolean
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
