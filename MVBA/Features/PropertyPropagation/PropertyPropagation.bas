' Module: PropertyPropagation
' Description: Propagates a trigger "label cell"'s full text as the value of a user-chosen custom
'              property onto every OTHER element of the cell's graphic group (cells or not). Opt-in,
'              OFF by default (ARES_Property_Propagation). Reuses the deferred change/idle pipeline:
'              capture in ElementChangeHandler.IChangeTrackEvents_ElementChanged, deferred processing
'              in ElementChangeHandler.ProcessElement (Depth 0). Custom-property attach/read/write/
'              detach go through CustomPropertyHandler; the item-type definitions live in the "ARES"
'              DGNLib (strategy A - VBA never authors item types).
'
'              Trigger condition (all must hold): the master switch is ON, the element IsCellElement,
'              its AsCellElement.Name is a member of the | -list ARES_Propagation_Cells (case-
'              insensitive, trimmed, non-empty), and the cell is in a real graphic group with at
'              least one other member. The value is the cell's full concatenated text
'              (StringsInEl.GetConcatenatedText). The target property is ARES_Propagation_Property
'              (a member of ARES_Custom_Property_List).
'
'              Loop-safety (load-bearing): AttachItem/SetPropertyValue/DetachItem write to the file
'              immediately, each firing a Modify event that the pipeline re-queues. Every write is
'              therefore COMPARE-GUARDED - never attach/set/detach when the sibling already holds the
'              intended state - so re-processing finds every value equal, writes nothing, emits no
'              event, and the cascade settles (the same convergence AutoLengths/color-sync rely on).
'
'              Deletion detaches the property from former siblings, DEFERRED: the Delete branch only
'              records intent (NoteDeletedTriggerCell, read-only Link.GetLink(BeforeChange)); the real
'              detach runs on idle in ProcessElement, reconciled (kept if another trigger cell still
'              governs the group) and present-guarded.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (global ARESConfig), CustomPropertyHandler, StringsInEl,
'               Link, LangManager, ErrorHandlerClass (global ErrorHandler)

Option Explicit

' Pending deferred detach: sibling element ID (DLongToString) -> target property name to detach.
' Recorded synchronously when a trigger cell is DELETED (read-only) and consumed on idle in
' ProcessElement. Keyed by the ID string so re-recording the same sibling is idempotent. An entry is
' only ever recorded with a NON-empty property name (ResolveTargetProperty succeeded).
Private moPendingDetach As Collection

' One-shot guards so the propagation statuses (PropagationValueRejected / PropagationPropertyInvalid /
' PropagationAttachFailed / PropagationMultipleTriggers) each surface only once per propagation batch.
' Reset at the start of each Propagate; the three fault ones also keep their English log on every
' occurrence. mbMultiShown is additionally reset in NoteDeletedTriggerCell (a pure-deletion idle batch
' runs no Propagate but must still surface the multi-trigger warning once).
Private mbRejectedShown As Boolean
Private mbInvalidShown As Boolean
Private mbAttachFailShown As Boolean
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
    If ARESConfig.ARES_PROPERTY_PROPAGATION Is Nothing Then Exit Function
    IsEnabled = CBool(ARESConfig.ARES_PROPERTY_PROPAGATION.Value)
    Exit Function

ErrorHandler:
    IsEnabled = False
End Function

' The 4-part trigger test: IsCellElement AND a real graphic group AND a non-empty cell-names list AND
' the cell name is a member of it (case-insensitive, trimmed). Mirrors the CellRedreaw.
' CheckInitialConditions name-vs-|-list precedent. Cheap IsCellElement test first (dormant fast path).
Public Function IsTriggerCell(ByVal oEl As element) As Boolean
    On Error GoTo ErrorHandler

    IsTriggerCell = False
    If oEl Is Nothing Then Exit Function
    If Not oEl.IsCellElement Then Exit Function
    If oEl.GraphicGroup = ARES_DEFAULT_GRAPHIC_GROUP_ID Then Exit Function

    Dim sList As String
    sList = PropagationCellsRaw()
    If Len(Trim(sList)) = 0 Then Exit Function

    Dim names() As String
    names = Split(sList, ARESConstants.ARES_VAR_DELIMITER)
    IsTriggerCell = NameInList(oEl.AsCellElement.Name, names)
    Exit Function

ErrorHandler:
    IsTriggerCell = False
End Function

' Depth-0 hook, called from ElementChangeHandler.ProcessElement (before the graphic-group filter).
' (1) If oEl was recorded as a former sibling of a deleted trigger cell, consume that entry and
'     reconcile against oEl's CURRENT group: re-propagate the surviving trigger cell's text if one
'     remains (round-2), else detach the property. (2) Otherwise, if oEl is itself a trigger cell,
'     propagate its text onto its siblings.
Public Sub ProcessElement(ByVal oEl As element)
    On Error GoTo ErrorHandler

    If oEl Is Nothing Then Exit Sub

    Dim sPendingProp As String
    sPendingProp = TakePendingProperty(oEl)
    If Len(sPendingProp) > 0 Then
        ' Reconcile the deferred detach against oEl's CURRENT group. Round-2: if a trigger cell still
        ' governs the group, RE-PROPAGATE its current text onto oEl (compare-guarded) instead of
        ' leaving the deleted cell's stale value; detach only when no trigger cell remains. The survivor
        ' is the first in scan order (deterministic -> convergent); >=2 survivors -> one-shot warning.
        Dim oSurvivor As element
        Dim bMulti As Boolean
        Set oSurvivor = FindGroupTriggerCell(oEl, bMulti)
        If oSurvivor Is Nothing Then
            CustomPropertyHandler.RemoveItemFromElement oEl, sPendingProp
        Else
            ApplyValueToSibling oEl, sPendingProp, StringsInEl.GetConcatenatedText(oSurvivor)
            If bMulti Then ReportMultipleTriggers
        End If
        Exit Sub
    End If

    If IsTriggerCell(oEl) Then Propagate oEl
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation.ProcessElement"
End Sub

' Delete-branch hook, called synchronously from IChangeTrackEvents_ElementChanged (READ-ONLY - never
' writes to the model here). If the deleted element was a trigger cell and the target property
' resolves, record each former sibling (els, already computed by ShouldQueueForDeletion via
' Link.GetLink(BeforeChange)) into the pending-detach set for a reconciled detach on idle.
Public Sub NoteDeletedTriggerCell(ByVal oDeletedCell As element, ByRef els() As element)
    On Error GoTo ErrorHandler

    If oDeletedCell Is Nothing Then Exit Sub
    If Not IsTriggerCell(oDeletedCell) Then Exit Sub

    ' Round-2: reset the multi-trigger one-shot here too. A pure-deletion idle batch consumes pending
    ' detaches (which may warn) but runs no Propagate, so this is the only reset point that covers it.
    mbMultiShown = False

    Dim P As String
    P = ResolveTargetProperty()
    If Len(P) = 0 Then Exit Sub

    If Not HasElements(els) Then Exit Sub

    EnsurePendingDetach

    Dim i As Long
    Dim s As element
    Dim sKey As String
    For i = LBound(els) To UBound(els)
        Set s = els(i)
        If Not s Is Nothing Then
            sKey = DLongToString(s.ID)
            If Not HasPending(sKey) Then moPendingDetach.Add P, sKey
        End If
    Next i
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation.NoteDeletedTriggerCell"
End Sub

'######################################################################################################################
'                                          PRIVATE HELPERS
'######################################################################################################################

' Resolve the target property, compute the cell's full text, and apply it (compare-guarded) to every
' graphic-group sibling. No sibling, or an empty/invalid property, means nothing to do.
Private Sub Propagate(ByVal oCell As element)
    On Error GoTo ErrorHandler

    mbRejectedShown = False
    mbInvalidShown = False
    mbAttachFailShown = False
    mbMultiShown = False

    Dim P As String
    P = ResolveTargetProperty()
    If Len(P) = 0 Then Exit Sub                 ' empty (silent) or invalid (already logged) -> nothing to write

    Dim sValue As String
    sValue = StringsInEl.GetConcatenatedText(oCell)

    Dim els() As element
    els = Link.GetLink(oCell)                   ' siblings in the same graphic group (excludes the cell)
    If Not HasElements(els) Then Exit Sub

    Dim i As Long
    Dim s As element
    Dim bMulti As Boolean
    bMulti = False
    For i = LBound(els) To UBound(els)
        Set s = els(i)
        If Not s Is Nothing Then
            ApplyValueToSibling s, P, sValue
            ' oCell is itself a trigger, so ANY trigger sibling means >=2 triggers in the group.
            ' Nested If (not And) - no short-circuit in VBA; also skips the test once already known.
            If Not bMulti Then
                If IsTriggerCell(s) Then bMulti = True
            End If
        End If
    Next i

    ' Round-2 (Demand 1): warn once when >=2 trigger cells share the group (last-modified wins).
    If bMulti Then ReportMultipleTriggers
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation.Propagate"
End Sub

' Compare-before-write on a single sibling (loop-safety). Non-empty value: attach if missing then
' set; set if a different value is present; no-op if already equal. Empty value: clear only a sibling
' that already carries a non-empty value; never newly attach an empty property. A rejected write
' (SetPropertyValueToElement False, e.g. a constrained picklist) is non-fatal: log + one-shot status.
' A failed ATTACH (AttachItemToElement False, e.g. the "ARES" DGNLib / ItemType not resolvable in the
' session) is likewise non-fatal but no longer silent: log + one-shot PropagationAttachFailed, then
' skip this sibling.
Private Sub ApplyValueToSibling(ByVal s As element, ByVal P As String, ByVal value As String)
    On Error GoTo ErrorHandler

    If s Is Nothing Then Exit Sub

    Dim vCurrent As Variant
    Dim bHasProp As Boolean
    Dim sCurrent As String

    vCurrent = CustomPropertyHandler.GetPropertyValueFromElement(s, P, P)
    bHasProp = Not IsNull(vCurrent)
    If bHasProp Then sCurrent = CStr(vCurrent) Else sCurrent = ""

    If Len(value) > 0 Then
        If Not bHasProp Then
            If Not CustomPropertyHandler.AttachItemToElement(s, P) Then
                ReportAttachFailed P
                Exit Sub
            End If
            If Not CustomPropertyHandler.SetPropertyValueToElement(s, P, value) Then ReportRejected
        ElseIf sCurrent <> value Then
            If Not CustomPropertyHandler.SetPropertyValueToElement(s, P, value) Then ReportRejected
        End If
        ' already equal -> no-op (loop-safety)
    Else
        ' Empty value: only clear a sibling that already carries a non-empty value.
        If bHasProp Then
            If Len(sCurrent) > 0 Then
                If Not CustomPropertyHandler.SetPropertyValueToElement(s, P, "") Then ReportRejected
            End If
        End If
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation.ApplyValueToSibling"
End Sub

' Trim the configured target property; "" (dormant, silent) when empty; validate membership in the
' managed custom-property list (strategy A). A non-member is not written: log (English) + one-shot
' PropagationPropertyInvalid status, and return "".
Private Function ResolveTargetProperty() As String
    On Error GoTo ErrorHandler

    ResolveTargetProperty = ""

    Dim sProp As String
    sProp = Trim(PropagationPropertyRaw())
    If Len(sProp) = 0 Then Exit Function

    Dim names() As String
    names = CustomPropertyHandler.GetCustomPropertyNames()
    If NameInList(sProp, names) Then
        ResolveTargetProperty = sProp
    Else
        ' Log every occurrence (English); surface the status only once per batch (symmetric with the
        ' rejected-write one-shot). mbInvalidShown is reset at the start of each Propagate.
        ErrorHandler.HandleError "Property propagation: target property '" & sProp & "' is not a member of ARES_Custom_Property_List", 0, "", "PropertyPropagation.ResolveTargetProperty"
        If Not mbInvalidShown Then
            LangManager.ShowStatusT "PropagationPropertyInvalid"
            mbInvalidShown = True
        End If
        ResolveTargetProperty = ""
    End If
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation.ResolveTargetProperty"
    ResolveTargetProperty = ""
End Function

' Round-2: reconcile helper for the deferred detach. Scans oEl's CURRENT graphic group (the whole
' group, including oEl itself via Link.GetLink ReturnMe:=True) and returns the FIRST trigger cell in
' scan order (Nothing if none remains). Sets bMultiple = True when >=2 trigger cells are present
' (drives the PropagationMultipleTriggers warning). Scan order is deterministic for a fixed group, so
' every consumed sibling resolves to the SAME survivor -> the re-propagation converges (Design Notes).
' Public (not Private as the Code Map listed) so PropertyPropagationTest can assert survivor + bMultiple
' directly - a Private helper is not callable from the separate UnitTesting module. Read-only query.
Public Function FindGroupTriggerCell(ByVal oEl As element, ByRef bMultiple As Boolean) As element
    On Error GoTo ErrorHandler

    Set FindGroupTriggerCell = Nothing
    bMultiple = False
    If oEl Is Nothing Then Exit Function
    If Not oEl.IsGraphical Then Exit Function
    If oEl.GraphicGroup = ARES_DEFAULT_GRAPHIC_GROUP_ID Then Exit Function

    Dim members() As element
    members = Link.GetLink(oEl, True)
    If Not HasElements(members) Then Exit Function

    Dim i As Long
    Dim nFound As Long
    nFound = 0
    For i = LBound(members) To UBound(members)
        If IsTriggerCell(members(i)) Then
            If FindGroupTriggerCell Is Nothing Then Set FindGroupTriggerCell = members(i)
            nFound = nFound + 1
            If nFound >= 2 Then
                bMultiple = True
                Exit Function
            End If
        End If
    Next i
    Exit Function

ErrorHandler:
    Set FindGroupTriggerCell = Nothing
    bMultiple = False
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

' Log the rejected write (English, Number 0) and surface PropagationValueRejected ONCE per batch.
Private Sub ReportRejected()
    On Error Resume Next
    ErrorHandler.HandleError "Property propagation: target property rejected the value", 0, "", "PropertyPropagation.ApplyValueToSibling"
    If Not mbRejectedShown Then
        LangManager.ShowStatusT "PropagationValueRejected"
        mbRejectedShown = True
    End If
End Sub

' Log the failed attach (English, Number 0, incl. property name) and surface PropagationAttachFailed
' ONCE per batch. A False from AttachItemToElement usually means the "ARES" DGNLib / the ItemType is
' not resolvable in the session (MS_DGNLIBLIST) - previously a fully silent no-op.
Private Sub ReportAttachFailed(ByVal P As String)
    On Error Resume Next
    ErrorHandler.HandleError "Property propagation: could not attach property '" & P & "' (ARES item-type library / item type not found)", 0, "", "PropertyPropagation.ApplyValueToSibling"
    If Not mbAttachFailShown Then
        LangManager.ShowStatusT "PropagationAttachFailed"
        mbAttachFailShown = True
    End If
End Sub

' Round-2 (Demand 1): surface PropagationMultipleTriggers ONCE per batch (deduped via mbMultiShown,
' reset in Propagate and NoteDeletedTriggerCell). This is USER FEEDBACK, not a fault: per the Design
' Note it is status-only and does NOT write an English .log line (unlike ReportRejected /
' ReportAttachFailed). Deviation from the Code Map wording ("mirror ReportRejected: English log +
' status") in favour of the authoritative Design Note.
Private Sub ReportMultipleTriggers()
    On Error Resume Next
    If Not mbMultiShown Then
        LangManager.ShowStatusT "PropagationMultipleTriggers"
        mbMultiShown = True
    End If
End Sub

' Raw ARES_Propagation_Cells ("" when unset). Lazily initialises ARESConfig like the other modules.
Private Function PropagationCellsRaw() As String
    On Error GoTo ErrorHandler

    PropagationCellsRaw = ""
    If ARESConfig Is Nothing Then Exit Function
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    If ARESConfig.ARES_PROPAGATION_CELLS Is Nothing Then Exit Function
    PropagationCellsRaw = ARESConfig.ARES_PROPAGATION_CELLS.Value
    Exit Function

ErrorHandler:
    PropagationCellsRaw = ""
End Function

' Raw ARES_Propagation_Property ("" when unset). Lazily initialises ARESConfig like the other modules.
Private Function PropagationPropertyRaw() As String
    On Error GoTo ErrorHandler

    PropagationPropertyRaw = ""
    If ARESConfig Is Nothing Then Exit Function
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    If ARESConfig.ARES_PROPAGATION_PROPERTY Is Nothing Then Exit Function
    PropagationPropertyRaw = ARESConfig.ARES_PROPAGATION_PROPERTY.Value
    Exit Function

ErrorHandler:
    PropagationPropertyRaw = ""
End Function

' Lazily create the pending-detach collection.
Private Sub EnsurePendingDetach()
    If moPendingDetach Is Nothing Then Set moPendingDetach = New Collection
End Sub

' True if sKey is already present in the pending-detach set.
Private Function HasPending(ByVal sKey As String) As Boolean
    On Error Resume Next
    HasPending = False
    If moPendingDetach Is Nothing Then Exit Function
    Err.Clear
    Dim v As Variant
    v = moPendingDetach(sKey)                   ' raises when the key is absent
    HasPending = (Err.Number = 0)
    Err.Clear
End Function

' If oEl.ID is in the pending-detach set, remove and RETURN the recorded property name; otherwise "".
' Entries always carry a non-empty property, so "" unambiguously means "not pending".
Private Function TakePendingProperty(ByVal oEl As element) As String
    On Error GoTo ErrorHandler

    TakePendingProperty = ""
    If moPendingDetach Is Nothing Then Exit Function
    If moPendingDetach.Count = 0 Then Exit Function

    Dim sKey As String
    sKey = DLongToString(oEl.ID)
    If Not HasPending(sKey) Then Exit Function

    TakePendingProperty = CStr(moPendingDetach(sKey))
    moPendingDetach.Remove sKey
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
