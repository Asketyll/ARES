' Module: PropertyCalculation
' Description: The VALUE-CALCULATION engine (redécoupage, epic 11). Writes a trigger "label cell"'s
'              full text as the VALUE of a target custom property onto every OTHER element of the
'              cell's graphic group - but ONLY where that property is ALREADY ATTACHED (by a
'              PropertyTagging rule). It NEVER attaches, and never calls CustomPropertyHandler
'              attach/detach directly. A member that does not carry the target property is SKIPPED
'              (the frontier: attach/detach is the tagger's domain). Opt-in, OFF by default
'              (ARES_Property_Calc). Reuses the deferred change/idle pipeline: capture in
'              ElementChangeHandler.IChangeTrackEvents_ElementChanged, deferred processing in
'              ElementChangeHandler.ProcessElement (Depth 0). Custom-property read/write go through
'              CustomPropertyHandler; item-type definitions live in the "ARES" DGNLib (strategy A).
'
'              PHASE-1 DORMANT (grammar v2, story 13-2): the @cell=prop value seam was removed, so this
'              engine no longer derives any trigger or target - IsTriggerCell returns False and
'              PropertyTagging.GetCellGroupProperties no longer exists. The value-WRITE machinery
'              (ApplyValueToSibling + the compare/transition loop-safety guards + IsDetachEmptyEnabled) is
'              retained BYTE-INTACT as scaffolding for phase 2 (per-property calculation rules), which will
'              re-derive the triggers/targets and re-wire this module. The master switch ARES_Property_Calc
'              is still read (IsEnabled / IsAnyFeatureEnabled), so the engine is enabled but INERT. The
'              frontier check for phase 2 is CustomPropertyHandler.IsItemAttachedToElement (HasItems, NOT
'              Null-inference: an attached-but-empty property also reads back Null).
'
'              Emptying semantics (round-4): when a value is emptied (trigger cell text emptied, or the
'              cell deleted with no surviving trigger cell), by default the value is CLEARED and the
'              property stays attached; with ARES_Calc_Detach_Empty ON, the detach is DELEGATED
'              to the tagger (PropertyTagging.DetachRuleProperty) instead - a governing rule then
'              re-attaches the property empty on the next pass (rules win). The detach fires ONLY on a
'              real non-empty -> empty TRANSITION (the member currently holds a non-empty value).
'
'              Loop-safety (load-bearing): SetPropertyValue / the delegated DetachItem write to the
'              file immediately, each firing a Modify event that the pipeline re-queues. Every value
'              write is COMPARE-GUARDED and the detach-on-empty is additionally TRANSITION-GUARDED
'              (Len(current) > 0) -> at most one detach per emptying: a rule re-attaches empty, the next
'              empty visit reads "" and no-ops, so the cascade settles (the convergence AutoLengths/
'              color-sync rely on). An un-guarded detach-on-empty would oscillate against a re-attaching
'              rule forever.
'
'              Deletion is DEFERRED + reconciled: the Delete branch only records intent
'              (NoteDeletedTriggerCell, read-only Link.GetLink(BeforeChange)); on idle ProcessElement
'              reconciles each former sibling against its CURRENT group - re-apply a surviving
'              trigger cell's text, else empty the value (clear or delegated-detach per the option).
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (global ARESConfig), CustomPropertyHandler, PropertyTagging,
'               StringsInEl, Link, LangManager, ErrorHandlerClass (global ErrorHandler)

Option Explicit

' Pending deferred clear/detach: former-sibling element ID (DLongToString) -> the deleted cell's target
' properties, a |-joined list (ARES_VAR_DELIMITER). Recorded synchronously when a trigger cell is DELETED
' (read-only) and consumed on idle in ProcessElement, where the consume path splits the list and
' reconciles EACH target property independently. Keyed by the ID string so re-recording the same sibling
' is idempotent. An entry always carries >=1 non-empty target. (Phase-1 dormant scaffolding.)
' Named "Clear" (not "Detach"): the value engine clears the value, or delegates a detach to the tagger
' with the option ON - it never detaches directly.
Private moPendingClear As Collection

' One-shot guards so the calculation statuses (CalculationValueRejected / CalculationNoTarget /
' CalculationMultipleTriggers) each surface only once per calculation batch. Reset at the start of each
' Calculate; the fault one (Rejected) also keeps its English log on every occurrence; NoTarget and
' Multiple are user feedback (status-only, no log). mbMultiShown is additionally reset in
' NoteDeletedTriggerCell (a pure-deletion idle batch runs no Calculate but must still surface the
' multi-trigger warning once).
Private mbRejectedShown As Boolean
Private mbNoTargetShown As Boolean
Private mbMultiShown As Boolean

'######################################################################################################################
'                                          PUBLIC SURFACE
'######################################################################################################################

' Master switch. Lazily initialises ARESConfig like the other feature modules.
Public Function IsEnabled() As Boolean
    On Error GoTo ErrorHandler

    IsEnabled = False
    If ARESConfig Is Nothing Then Exit Function
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    If ARESConfig.ARES_PROPERTY_CALC Is Nothing Then Exit Function
    IsEnabled = CBool(ARESConfig.ARES_PROPERTY_CALC.Value)
    Exit Function

ErrorHandler:
    IsEnabled = False
End Function

' Trigger test - PHASE-1 DORMANT (grammar v2, story 13-2). Grammar v2 removed the @cell=prop value seam,
' so nothing is a trigger any more: this returns False unconditionally. The signature is kept so callers
' stay compilable; phase 2 re-introduces the trigger derivation (from per-property calculation rules).
Public Function IsTriggerCell(ByVal oEl As element) As Boolean
    On Error GoTo ErrorHandler

    IsTriggerCell = False
    Exit Function

ErrorHandler:
    IsTriggerCell = False
End Function

' Depth-0 hook, called from ElementChangeHandler.ProcessElement (before the graphic-group filter).
' (1) If oEl was recorded as a former sibling of a deleted trigger cell, consume that entry and
'     reconcile against oEl's CURRENT group: re-apply the surviving trigger cell's text if one
'     remains (round-2), else empty the value (clear, or delegated tagger-detach with the option ON).
'     (2) Otherwise, if oEl is itself a trigger cell, write its text onto its siblings.
Public Sub ProcessElement(ByVal oEl As element)
    On Error GoTo ErrorHandler

    If oEl Is Nothing Then Exit Sub

    Dim sPending As String
    sPending = TakePendingProperty(oEl)
    If Len(sPending) > 0 Then
        ' The pending value is a |-joined list of the deleted cell's target properties (12-1). Reconcile
        ' EACH target independently against oEl's CURRENT group: if a surviving trigger cell whose targets
        ' include P still governs the group, RE-APPLY its text into P (compare-guarded); else EMPTY it
        ' via ApplyValueToSibling(oEl, P, "") - clears (option OFF) or delegates a tagger-detach (option
        ' ON), transition-guarded. The per-property survivor is the first in scan order (deterministic ->
        ' convergent); >=2 survivors targeting P -> one-shot warning. NO direct CustomPropertyHandler detach.
        Dim pend() As String
        Dim k As Long
        Dim oSurvivor As element
        Dim bMulti As Boolean
        pend = Split(sPending, ARESConstants.ARES_VAR_DELIMITER)
        For k = LBound(pend) To UBound(pend)
            If Len(pend(k)) > 0 Then
                Set oSurvivor = FindGroupTriggerCellForProperty(oEl, pend(k), bMulti)
                If oSurvivor Is Nothing Then
                    ApplyValueToSibling oEl, pend(k), ""
                Else
                    ApplyValueToSibling oEl, pend(k), StringsInEl.GetConcatenatedText(oSurvivor)
                    If bMulti Then ReportMultipleTriggers
                End If
            End If
        Next k
        Exit Sub
    End If

    If IsTriggerCell(oEl) Then Calculate oEl
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation.ProcessElement"
End Sub

' Delete-branch hook, called synchronously from IChangeTrackEvents_ElementChanged (READ-ONLY - never
' writes to the model here). If the deleted element was a @cell trigger, resolve its target properties
' (the former @cell target derivation, now removed - phase-1 dormant) and record, per former
' sibling (els, already computed by ShouldQueueForDeletion via Link.GetLink(BeforeChange)), the |-joined
' target list into the pending-clear set for a per-property reconciled emptying on idle.
Public Sub NoteDeletedTriggerCell(ByVal oDeletedCell As element, ByRef els() As element)
    On Error GoTo ErrorHandler

    If oDeletedCell Is Nothing Then Exit Sub
    If Not IsTriggerCell(oDeletedCell) Then Exit Sub

    ' Round-2: reset the multi-trigger one-shot here too. A pure-deletion idle batch consumes pending
    ' entries (which may warn) but runs no Calculate, so this is the only reset point that covers it.
    mbMultiShown = False

    ' Phase-1 dormant: no target derivation (GetCellGroupProperties is gone). Unreachable anyway - the
    ' IsTriggerCell guard above always exits first - but kept compilable with the empty convention.
    Dim props() As String
    ReDim props(0 To 0)
    props(0) = ""
    If Len(props(LBound(props))) = 0 Then Exit Sub

    ' All entries are non-empty past the guard above (the empty convention is a single ""); join them.
    Dim sJoined As String
    sJoined = Join(props, ARESConstants.ARES_VAR_DELIMITER)

    If Not HasElements(els) Then Exit Sub

    EnsurePendingClear

    Dim i As Long
    Dim s As element
    Dim sKey As String
    For i = LBound(els) To UBound(els)
        Set s = els(i)
        If Not s Is Nothing Then
            sKey = DLongToString(s.ID)
            If Not HasPending(sKey) Then moPendingClear.Add sJoined, sKey
        End If
    Next i
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation.NoteDeletedTriggerCell"
End Sub

'######################################################################################################################
'                                          PRIVATE HELPERS
'######################################################################################################################

' Resolve the cell's target properties from the @cell rules, compute its full text once, and write it
' (compare-guarded, frontier-checked) into EACH target on every graphic-group sibling carrying it. No
' sibling, or no @cell target, means nothing to do.
Private Sub Calculate(ByVal oCell As element)
    On Error GoTo ErrorHandler

    mbRejectedShown = False
    mbNoTargetShown = False
    mbMultiShown = False

    ' Phase-1 dormant: no target derivation (GetCellGroupProperties is gone). Unreachable anyway - Calculate
    ' is only called behind IsTriggerCell (always False) - but kept compilable with the empty convention.
    Dim props() As String
    ReDim props(0 To 0)
    props(0) = ""
    If Len(props(LBound(props))) = 0 Then Exit Sub    ' phase-1 dormant -> nothing to write

    Dim sValue As String
    sValue = StringsInEl.GetConcatenatedText(oCell)

    Dim els() As element
    els = Link.GetLink(oCell)                   ' siblings in the same graphic group (excludes the cell)
    If Not HasElements(els) Then Exit Sub

    Dim i As Long, j As Long
    Dim s As element
    Dim bMulti As Boolean
    Dim nAttached As Long
    bMulti = False
    nAttached = 0
    For i = LBound(els) To UBound(els)
        Set s = els(i)
        If Not s Is Nothing Then
            ' Write the cell's text into EACH target property (per-rule targets, 12-1). Each write is
            ' independently compare-guarded/frontier-checked; ApplyValueToSibling returns True when s
            ' already carries that target - count the (sibling,target) hits to detect "none carry it".
            For j = LBound(props) To UBound(props)
                If Len(props(j)) > 0 Then
                    If ApplyValueToSibling(s, props(j), sValue) Then nAttached = nAttached + 1
                End If
            Next j
            ' Multi-trigger is scoped to a SHARED target (12-1): a trigger sibling conflicts only if its
            ' targets intersect this cell's. Nested If (not And) - no short-circuit; skips once known.
            If Not bMulti Then
                If IsTriggerCell(s) Then
                    ' Phase-1 dormant + unreachable (IsTriggerCell always False); GetCellGroupProperties gone.
                    If TargetsIntersect(props, props) Then bMulti = True
                End If
            End If
        End If
    Next i

    ' Discoverability (residual guard): siblings exist but NONE carried any target -> the attach did not
    ' happen (Property Tagging OFF, or the DGNLib/ItemType unresolved). One-shot hint (status-only, no log).
    If nAttached = 0 Then ReportNoTarget

    ' Multi-trigger (shared target): warn once (last-modified wins).
    If bMulti Then ReportMultipleTriggers
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation.Calculate"
End Sub

' The frontier + compare-before-write on a single sibling (loop-safety). Returns True when s ALREADY
' carries the target property P (whether or not a write happened) - the caller counts these to detect
' the "no member carries P" misconfiguration. The value engine NEVER attaches and never calls
' CustomPropertyHandler detach directly:
'   - P not attached (IsItemAttachedToElement False) -> SKIP (return False). Attach is the tagger's job.
'   - non-empty value, different from current       -> set (compare-guarded); rejection -> one-shot status.
'   - non-empty value, equal to current             -> no-op (loop-safety).
'   - empty value, current non-empty (a real emptying TRANSITION):
'         option OFF -> clear the value ("");  option ON -> delegate a detach to the tagger
'         (PropertyTagging.DetachRuleProperty). This is the ONLY detach path, gated on BOTH the option
'         AND the non-empty->empty transition (the load-bearing loop-safety guard).
'   - empty value, current already empty            -> no-op (transition guard: no re-detach).
Private Function ApplyValueToSibling(ByVal s As element, ByVal P As String, ByVal value As String) As Boolean
    On Error GoTo ErrorHandler

    ApplyValueToSibling = False
    If s Is Nothing Then Exit Function

    ' Frontier: write only where P is ALREADY attached (HasItems, not Null-inference - an attached-but-
    ' empty property also reads back Null). Not attached -> skip; attach stays the tagger's domain.
    If Not CustomPropertyHandler.IsItemAttachedToElement(s, P) Then Exit Function
    ApplyValueToSibling = True

    ' Read the current value. Nested read-then-branch keeps CStr off a possible array (no short-circuit
    ' in VBA); an attached-but-empty property reads back Null -> sCurrent "".
    Dim vCurrent As Variant
    Dim sCurrent As String
    vCurrent = CustomPropertyHandler.GetPropertyValueFromElement(s, P, P)
    If IsNull(vCurrent) Then sCurrent = "" Else sCurrent = CStr(vCurrent)

    If Len(value) > 0 Then
        ' Non-empty value: set only when different (compare-guarded).
        If sCurrent <> value Then
            If Not CustomPropertyHandler.SetPropertyValueToElement(s, P, value) Then ReportRejected
        End If
        ' already equal -> no-op (loop-safety)
    Else
        ' Empty value: act ONLY on a real non-empty -> empty TRANSITION (an already-empty property is a
        ' no-op, so a rule that re-attaches P empty does not re-trigger a detach - this makes ON terminate).
        If Len(sCurrent) > 0 Then
            If IsDetachEmptyEnabled() Then
                ' Option ON: delegate the detach to the tagger (the only permitted detach path).
                PropertyTagging.DetachRuleProperty s, P
            Else
                ' Option OFF: clear the value; the property stays attached.
                If Not CustomPropertyHandler.SetPropertyValueToElement(s, P, "") Then ReportRejected
            End If
        End If
    End If
    Exit Function

ErrorHandler:
    ' A fault mid-write does not un-attach P; the return value (set True once past the frontier) is only
    ' used to detect "no member carried P", so leaving it as-is is correct.
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation.ApplyValueToSibling"
End Function

' Round-4 option (ARES_Calc_Detach_Empty): when True, an emptied value is DETACHED (delegated to
' the tagger) instead of cleared. Mirrors IsEnabled - fail-closed False on any nil; lazy ARESConfig init.
Private Function IsDetachEmptyEnabled() As Boolean
    On Error GoTo ErrorHandler

    IsDetachEmptyEnabled = False
    If ARESConfig Is Nothing Then Exit Function
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    If ARESConfig.ARES_CALC_DETACH_EMPTY Is Nothing Then Exit Function
    IsDetachEmptyEnabled = CBool(ARESConfig.ARES_CALC_DETACH_EMPTY.Value)
    Exit Function

ErrorHandler:
    IsDetachEmptyEnabled = False
End Function

' Reconcile helper (12-1, per-property): scans oEl's CURRENT graphic group (the whole group, including
' oEl itself via Link.GetLink ReturnMe:=True) and returns the FIRST trigger cell in scan order whose
' target properties INCLUDE P (Nothing if none remains). Sets bMultiple = True when >=2 such cells
' target P (drives the CalculationMultipleTriggers warning). Scan order is deterministic, so every
' consumed sibling resolves to the SAME survivor for P -> the re-calculation converges (Design Notes).
' Public so PropertyCalculationTest can assert survivor + bMultiple directly. Read-only query.
Public Function FindGroupTriggerCellForProperty(ByVal oEl As element, ByVal P As String, ByRef bMultiple As Boolean) As element
    On Error GoTo ErrorHandler

    Set FindGroupTriggerCellForProperty = Nothing
    bMultiple = False
    If oEl Is Nothing Then Exit Function
    If Not oEl.IsGraphical Then Exit Function
    If oEl.GraphicGroup = ARES_DEFAULT_GRAPHIC_GROUP_ID Then Exit Function

    Dim members() As element
    members = Link.GetLink(oEl, True)
    If Not HasElements(members) Then Exit Function

    Dim i As Long
    Dim nFound As Long
    Dim mProps() As String
    nFound = 0
    For i = LBound(members) To UBound(members)
        If IsTriggerCell(members(i)) Then
            ' Phase-1 dormant + unreachable (IsTriggerCell always False); GetCellGroupProperties gone.
            ReDim mProps(0 To 0)
            mProps(0) = ""
            If NameInList(P, mProps) Then
                If FindGroupTriggerCellForProperty Is Nothing Then Set FindGroupTriggerCellForProperty = members(i)
                nFound = nFound + 1
                If nFound >= 2 Then
                    bMultiple = True
                    Exit Function
                End If
            End If
        End If
    Next i
    Exit Function

ErrorHandler:
    Set FindGroupTriggerCellForProperty = Nothing
    bMultiple = False
End Function

' True when any non-empty target in a() also appears in b() (case-insensitive). Scopes the multi-trigger
' warning to a SHARED target property: two @cell trigger cells conflict only when their target sets
' intersect (different-property cells are a valid, silent configuration).
Private Function TargetsIntersect(ByRef a() As String, ByRef b() As String) As Boolean
    On Error GoTo ErrorHandler

    TargetsIntersect = False
    Dim i As Long
    For i = LBound(a) To UBound(a)
        If Len(a(i)) > 0 Then
            If NameInList(a(i), b) Then
                TargetsIntersect = True
                Exit Function
            End If
        End If
    Next i
    Exit Function

ErrorHandler:
    TargetsIntersect = False
End Function

' True when sName (trimmed) matches any member of names (trimmed), case-insensitive. Mirrors the
' CellRedreaw name-vs-|-list comparison but with StrComp/vbTextCompare + Trim (per the story).
Private Function NameInList(ByVal sName As String, ByRef names() As String) As Boolean
    On Error GoTo ErrorHandler

    NameInList = False
    Dim sTarget As String
    sTarget = Trim(sName)
    If Len(sTarget) = 0 Then Exit Function

    Dim i As Long
    For i = LBound(names) To UBound(names)
        If StrComp(Trim(names(i)), sTarget, vbTextCompare) = 0 Then
            NameInList = True
            Exit Function
        End If
    Next i
    Exit Function

ErrorHandler:
    NameInList = False
End Function

' Log the rejected write (English, Number 0) and surface CalculationValueRejected ONCE per batch.
Private Sub ReportRejected()
    On Error Resume Next
    ErrorHandler.HandleError "Property calculation: target property rejected the value", 0, "", "PropertyCalculation.ApplyValueToSibling"
    If Not mbRejectedShown Then
        LangManager.ShowStatusT "CalculationValueRejected"
        mbRejectedShown = True
    End If
End Sub

' Surface CalculationNoTarget ONCE per batch: a calculation ran with siblings present but NONE carried
' the target property (the value engine writes only where a rule already attached P). USER FEEDBACK,
' not a fault - status-only, no English .log (like ReportMultipleTriggers). Hints the user to add an
' attach rule in Property Tagging (GUI 1).
Private Sub ReportNoTarget()
    On Error Resume Next
    If Not mbNoTargetShown Then
        LangManager.ShowStatusT "CalculationNoTarget"
        mbNoTargetShown = True
    End If
End Sub

' Round-2 (Demand 1): surface CalculationMultipleTriggers ONCE per batch (deduped via mbMultiShown,
' reset in Calculate and NoteDeletedTriggerCell). This is USER FEEDBACK, not a fault: per the Design
' Note it is status-only and does NOT write an English .log line (unlike ReportRejected). Deviation
' from the Code Map wording ("mirror ReportRejected: English log + status") in favour of the
' authoritative Design Note.
Private Sub ReportMultipleTriggers()
    On Error Resume Next
    If Not mbMultiShown Then
        LangManager.ShowStatusT "CalculationMultipleTriggers"
        mbMultiShown = True
    End If
End Sub

' Lazily create the pending-clear collection.
Private Sub EnsurePendingClear()
    If moPendingClear Is Nothing Then Set moPendingClear = New Collection
End Sub

' True if sKey is already present in the pending-clear set.
Private Function HasPending(ByVal sKey As String) As Boolean
    On Error Resume Next
    HasPending = False
    If moPendingClear Is Nothing Then Exit Function
    Err.Clear
    Dim v As Variant
    v = moPendingClear(sKey)                    ' raises when the key is absent
    HasPending = (Err.Number = 0)
    Err.Clear
End Function

' If oEl.ID is in the pending-clear set, remove and RETURN the recorded property name; otherwise "".
' Entries always carry a non-empty property, so "" unambiguously means "not pending".
Private Function TakePendingProperty(ByVal oEl As element) As String
    On Error GoTo ErrorHandler

    TakePendingProperty = ""
    If moPendingClear Is Nothing Then Exit Function
    If moPendingClear.Count = 0 Then Exit Function

    Dim sKey As String
    sKey = DLongToString(oEl.ID)
    If Not HasPending(sKey) Then Exit Function

    TakePendingProperty = CStr(moPendingClear(sKey))
    moPendingClear.Remove sKey
    Exit Function

ErrorHandler:
    TakePendingProperty = ""
End Function

' Safe "array has at least one element" check (mirrors ElementChangeHandler.HasElements). UBound
' returns -1 for an empty array and raises for an uninitialised one.
Private Function HasElements(ByRef arr() As element) As Boolean
    On Error Resume Next
    HasElements = False
    If UBound(arr) <> -1 Then HasElements = True
    On Error GoTo 0
End Function
