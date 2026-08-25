' Class Module: ElementChangeHandler
' Description: Handles element change events for property tagging, calculation, actuation and rendering.
'              Implements IChangeTrackEvents for add/modify/delete, with frequency-based bulk-operation
'              detection.
'              `ARES_DIRECT_PROCESSING` (default True): processes each element synchronously inline
'              instead of queuing it for idle (False = older deferred fallback). Risk: a .Rewrite inside
'              ProcessElement can make MicroStation reenter IChangeTrackEvents_ElementChanged before the
'              outer call returns - mbProcessingDirect guards this by deferring the reentrant call instead.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, ARESConstants, LangManager, StringsInEl,
'               ElementInProcesseClass, ErrorHandlerClass, ReRegisterIdleHandler, PropertyTagging,
'               PropertyCalculation, PropertyActuator, PropertyRendering, CallStackClass

Option Explicit
Implements IChangeTrackEvents

' === PRIVATE MEMBERS ===
Private moIdleHandler As IdleEventHandler       ' Single persistent idle handler instance - created once, reused; only add/removed from the idle list
Private mbIdleHandlerRegistered As Boolean      ' Flag to prevent multiple idle handler registrations
Private moReRegHandler As ReRegisterIdleHandler ' Member-scoped ref to prevent premature VBA GC after bulk suspend
Private mbProcessingDirect As Boolean           ' Reentrancy guard for ARES_DIRECT_PROCESSING mode - see class header
Private mbReentrancyLogged As Boolean           ' One-shot flag: log the FIRST time ProcessDirect skips a reentrant call

' === BULK OPERATION DETECTION (Frequency-based) ===
' Detects bulk operations by tracking call frequency within a time window
Private mdEntryTime As Double                   ' Timestamp of first call in current window
Private mlCallCount As Long                     ' Counter for calls within current window
Private mlCachedThreshold As Long               ' Cached threshold — refreshed on window reset
Private mlCachedInterval As Long                ' Cached interval  — refreshed on window reset

' ========================================
' ICHANGETRACKEVENTS INTERFACE IMPLEMENTATION
' ========================================

' Not implemented - no undo/redo handling needed.
Private Sub IChangeTrackEvents_BeginUndoRedo(ByVal AfterUndoRedo As element, ByVal BeforeUndoRedo As element, ByVal Action As MsdChangeTrackAction, ByVal IsUndo As Boolean)
    On Error GoTo ErrorHandler
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.IChangeTrackEvents_BeginUndoRedo"
    Err.Clear
End Sub

' Main entry point for element add/modify/delete. AfterChange is Nothing on Delete; BeforeChange is
' Nothing on Add.
Private Sub IChangeTrackEvents_ElementChanged(ByVal AfterChange As element, ByVal BeforeChange As element, ByVal Action As MsdChangeTrackAction, CantBeUndone As Boolean)
    On Error GoTo ErrorHandler
    Dim els() As element
    Dim i As Long
    Dim bStackPushed As Boolean      ' True once Push below has run - guards Pop against a fault above
                                      ' it, so a reentrant call never pops an OUTER call's live frame

    ' PERFORMANCE: Detect and suspend during bulk operations
    If DetectAndSuspendBulkOperation() Then Exit Sub

    ' GUARD: Config not yet initialized (e.g. seed elements added before first idle fires)
    If Not ARESConfig.IsInitialized Then Exit Sub

    ' PERFORMANCE: Early exit if all ARES features are disabled
    If Not IsAnyFeatureEnabled() Then Exit Sub

    ' AfterChange/BeforeChange are mutually exclusive by construction of the Select Case below - pass
    ' whichever is set.
    If Not AfterChange Is Nothing Then
        CallStack.Push "ElementChangeHandler.IChangeTrackEvents_ElementChanged", AfterChange
    Else
        CallStack.Push "ElementChangeHandler.IChangeTrackEvents_ElementChanged", BeforeChange
    End If
    bStackPushed = True

    Select Case Action
        Case msdChangeTrackActionAdd
            If Not AfterChange Is Nothing Then
                If ShouldQueueElement(AfterChange) Then
                    QueueOrProcessDirect AfterChange
                End If
            End If

        Case msdChangeTrackActionModify
            If Not AfterChange Is Nothing Then
                If ShouldQueueElement(AfterChange) Then
                    QueueOrProcessDirect AfterChange
                End If
            End If

        Case msdChangeTrackActionDelete
            If Not BeforeChange Is Nothing Then
                If ShouldQueueForDeletion(BeforeChange, els) Then
                    ' Property Calculation needs no deletion hook here: deleting a trigger cell already
                    ' re-queues its former group members (via Link.GetLink), so their Depth-0 BEARING
                    ' pass re-resolves each carried calc property against the CURRENT group.
                    For i = LBound(els) To UBound(els)
                        QueueOrProcessDirect els(i)
                    Next i
                End If
            End If
    End Select
    CallStack.Pop
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.IChangeTrackEvents_ElementChanged"
    If bStackPushed Then CallStack.Pop
    Err.Clear
End Sub

' Not implemented - no post-undo/redo handling needed.
Private Sub IChangeTrackEvents_FinishUndoRedo(ByVal IsUndo As Boolean)
    On Error GoTo ErrorHandler
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.IChangeTrackEvents_FinishUndoRedo"
    Err.Clear
End Sub

' Not implemented - no change-mark handling needed.
Private Sub IChangeTrackEvents_Mark()
    On Error GoTo ErrorHandler
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.IChangeTrackEvents_Mark"
    Err.Clear
End Sub

' ========================================
' PRIVATE METHODS - Direct vs Deferred Dispatch
' ========================================

' Routes one already-filtered element to the deferred queue (default) or, under
' ARES_DIRECT_PROCESSING, straight to ProcessDirect - see class header.
Private Sub QueueOrProcessDirect(ByVal oElement As element)
    If CBool(ARESConfig.ARES_DIRECT_PROCESSING.Value) Then
        ProcessDirect oElement
    Else
        If ElementInProcesse.Add(oElement) Then EnsureIdleHandlerRegistered
    End If
End Sub

' Processes one element synchronously, inline, from within IChangeTrackEvents_ElementChanged
' (ARES_DIRECT_PROCESSING). mbProcessingDirect is the reentrancy guard - see class header for why a
' reentrant call falls back to the deferred queue instead of processing synchronously.
Private Sub ProcessDirect(ByVal oElement As element)
    On Error GoTo ErrorHandler

    If mbProcessingDirect Then
        If Not mbReentrancyLogged Then
            ErrorHandler.HandleError "ARES_Direct_Processing: reentrant ElementChanged call deferred to next idle (not processed synchronously by this call)", 0, "", "ElementChangeHandler.ProcessDirect"
            mbReentrancyLogged = True
        End If
        If ElementInProcesse.Add(oElement) Then EnsureIdleHandlerRegistered
        Exit Sub
    End If
    mbProcessingDirect = True

    ProcessElement oElement

    mbProcessingDirect = False
    Exit Sub

ErrorHandler:
    mbProcessingDirect = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.ProcessDirect"
    Err.Clear
End Sub

' ========================================
' PRIVATE METHODS - Idle Handler Management
' ========================================

' Ensures the single persistent IdleEventHandler is registered in MicroStation's idle list. Created once
' (lazily) and reused for the lifetime of this ChangeHandler - never destroyed/recreated per batch, which
' avoids ObjPtr reuse and stale-handler desync.
Private Sub EnsureIdleHandlerRegistered()
    On Error GoTo ErrorHandler

    If mbIdleHandlerRegistered Then Exit Sub
    If moIdleHandler Is Nothing Then Set moIdleHandler = New IdleEventHandler

    AddEnterIdleEventHandler moIdleHandler
    mbIdleHandlerRegistered = True

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.EnsureIdleHandlerRegistered"
End Sub

' Called by ReRegisterIdleHandler after bulk resume to release the member reference
Public Sub ClearReRegHandler()
    Set moReRegHandler = Nothing
End Sub

' Called by IdleEventHandler.CleanupHandler after it removes itself from the idle list; clears the
' "registered" flag so the next batch re-adds the (kept-alive) handler.
Public Sub ResetIdleHandlerFlag()
    On Error GoTo ErrorHandler

    mbIdleHandlerRegistered = False

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.ResetIdleHandlerFlag"
End Sub

' True when oHandler is the active persistent idle handler - lets IdleEventHandler.CleanupHandler tell
' the live handler apart from an orphan (boot's one-shot handler, or a stale instance after ChangeHandler
' was replaced) so only the live one can clear the queue / reset the registration flag.
Public Function IsActiveIdleHandler(ByVal oHandler As IdleEventHandler) As Boolean
    IsActiveIdleHandler = mbIdleHandlerRegistered And (Not moIdleHandler Is Nothing) And (oHandler Is moIdleHandler)
End Function

' ========================================
' PUBLIC METHODS - Element Processing
' ========================================

' Called by IdleEventHandler for each pending element. At Depth 0: property tagging, then calculation,
' then actuation, then rendering. Then Branch 1 (text/cell: nothing left to do) or Branch 2 (geometric:
' find linked text elements and recurse).
Public Sub ProcessElement(ByVal oElement As element, Optional ByVal Depth As Integer = 0)
    On Error GoTo ErrorHandler

    Const MAX_DEPTH As Integer = 20
    If Depth > MAX_DEPTH Then Exit Sub

    Dim i As Long                   ' Loop counter (Branch 2)
    Dim els() As element            ' Array of linked elements
    Dim linkedElement As element    ' Current linked element being processed

    ' Elements can become invalid if deleted between queueing and processing
    If oElement Is Nothing Then Exit Sub
    If Not IsElementValid(oElement) Then Exit Sub

    ' Pushed only past the cheap MAX_DEPTH/validity guards above; every Exit Sub below this point
    ' (5 of them) plus ErrorHandler pops this frame - see ErrorHandler for the reentrancy note.
    Dim bStackPushed As Boolean
    CallStack.Push "ElementChangeHandler.ProcessElement", oElement
    bStackPushed = True

    ' Depth-0 only, BEFORE the graphic-group filter below (so ungrouped elements are reached too):
    ' Tagging attaches -> Calculation fills -> Actuator paints -> Rendering writes the text. Order
    ' matters - each stage reads what the previous one just wrote, in the SAME pass.
    If Depth = 0 Then
        If CBool(ARESConfig.ARES_AUTO_PROPERTIES.Value) Then
            PropertyTagging.ApplyPropertyRules oElement
        End If
        If PropertyCalculation.IsEnabled Then
            PropertyCalculation.ProcessElement oElement
        End If
        If PropertyActuator.IsEnabled Then
            PropertyActuator.ProcessElement oElement
        End If
        If PropertyRendering.IsEnabled Then
            PropertyRendering.ProcessElement oElement
        End If
    End If

    ' Skip elements with the default graphic group (not part of ARES's group workflow)
    If oElement.GraphicGroup = ARES_DEFAULT_GRAPHIC_GROUP_ID Then
        CallStack.Pop
        Exit Sub
    End If

    ' Text/cell elements have nothing left to do beyond the Depth-0 hooks above, and are not geometric
    ' so Branch 2 does not apply either.
    If IsTextOrCellElement(oElement) Then
        CallStack.Pop
        Exit Sub
    End If

    ' ========================================
    ' BRANCH 2: GEOMETRIC ELEMENTS
    ' ========================================
    ' Find linked text elements and trigger their update. Also covers GroupLength/GroupColor rules,
    ' which have no name-pattern "trigger element" - re-evaluating every linked sibling here IS their
    ' only propagation mechanism.
    Dim bCalcEnabled As Boolean
    bCalcEnabled = PropertyCalculation.IsEnabled
    If Not (bCalcEnabled And (PropertyCalculation.HasGroupLengthRules Or PropertyCalculation.HasGroupColorRules)) Then
        CallStack.Pop
        Exit Sub
    End If

    els = Link.GetLink(oElement)
    If Not HasElements(els) Then
        CallStack.Pop
        Exit Sub
    End If

    For i = LBound(els) To UBound(els)
        Set linkedElement = els(i)
        If Not linkedElement Is Nothing Then
            If ElementInProcesse.Add(linkedElement) Then
                ProcessElement linkedElement, Depth + 1
            End If
            ' Depth > 0 skips the Depth-0-only calc hook, and the queue-add is wiped before a fresh pass
            ' reaches this sibling - a GroupLength/GroupColor-fed sibling needs this INLINE call instead.
            If bCalcEnabled Then PropertyCalculation.ProcessElement linkedElement
        End If
    Next i

    ' Drain the renderer's repaint hop AFTER the inline calc loop above (which can reach
    ' ApplyValueToSibling -> PropertyRendering.NoteDirtyGroup, after this element's own Depth-0 render
    ' hook already ran) so nothing is left stale or leaks across batches.
    If PropertyRendering.IsEnabled Then PropertyRendering.DrainRepaintHop

    CallStack.Pop
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.ProcessElement"
    If bStackPushed Then CallStack.Pop
    Err.Clear
End Sub

' ========================================
' PRIVATE HELPER FUNCTIONS
' ========================================

' True for Cell/Text/TextNode elements.
Private Function IsTextOrCellElement(ByVal element As element) As Boolean
    IsTextOrCellElement = (element.IsCellElement Or element.IsTextElement Or element.IsTextNodeElement)
End Function

' True if oElement is still valid/accessible (can become invalid if deleted between queueing and
' processing) and graphical.
Private Function IsElementValid(ByVal oElement As element) As Boolean
    On Error GoTo InvalidElement
    Dim nTest As Long

    IsElementValid = False
    nTest = oElement.Type    ' fails if the element is invalid/deleted
    If Not oElement.IsGraphical Then Exit Function

    IsElementValid = True
    Exit Function

InvalidElement:
    IsElementValid = False
End Function

' NOTE: IsNumericText lives in StringsInEl - this module no longer calls it. It is the NUMERIC GATE of
' PropertyRendering's forge protection; PropertyRendering is its only caller.

' ========================================
' PRIVATE HELPER FUNCTIONS - Bulk Operation Detection
' ========================================

' Frequency-based bulk-operation detection: ARES_Bulk_Threshold calls within ARES_Bulk_Interval ms ->
' suspend. True if a bulk operation was detected and the handler was suspended.
Private Function DetectAndSuspendBulkOperation() As Boolean
    On Error GoTo ErrorHandler
    Dim dCurrentTime As Double
    Dim dElapsedMs As Double

    DetectAndSuspendBulkOperation = False

    ' ARES's own element modifications during idle processing would trigger false bulk detection.
    If IsIdleProcessingActive() Then Exit Function

    ' Mirrors the guard above for ARES_DIRECT_PROCESSING: a group's own Calculation/Actuator/Rendering
    ' cascade across several siblings from one ProcessDirect call must not be miscounted as bulk.
    If mbProcessingDirect Then Exit Function

    ' Already suspended: in-flight events keep arriving during a tight bulk loop - skip them without
    ' re-running SuspendAndScheduleResume (would orphan the pending resume, leaving tracking dead).
    If IsChangeTrackingSuspended() Then
        If Not moReRegHandler Is Nothing Then
            DetectAndSuspendBulkOperation = True
            Exit Function
        End If
        ' Self-heal: suspended flag set but no resume pending - an orphaned suspension (ChangeHandler
        ' replaced on file open). Still receiving events means a live handler IS registered: clear the
        ' stale flag and fall through instead of skipping forever.
        ResetSuspensionState
    End If

    dCurrentTime = Timer

    ' Handle midnight rollover - Timer resets to 0 at midnight
    If dCurrentTime < mdEntryTime Then
        mlCachedThreshold = CLng(ARESConfig.ARES_BULK_THRESHOLD.Value)
        mlCachedInterval = CLng(ARESConfig.ARES_BULK_INTERVAL.Value)
        mdEntryTime = dCurrentTime
        mlCallCount = 1
        Exit Function
    End If

    dElapsedMs = (dCurrentTime - mdEntryTime) * 1000

    ' Outside the time window - reset and refresh cache
    If dElapsedMs > mlCachedInterval Or mlCachedInterval = 0 Then
        mlCachedThreshold = CLng(ARESConfig.ARES_BULK_THRESHOLD.Value)
        mlCachedInterval = CLng(ARESConfig.ARES_BULK_INTERVAL.Value)
        mdEntryTime = dCurrentTime
        mlCallCount = 1
        Exit Function
    End If

    mlCallCount = mlCallCount + 1

    If mlCallCount >= mlCachedThreshold Then
        SuspendAndScheduleResume
        DetectAndSuspendBulkOperation = True
    End If

    Exit Function

ErrorHandler:
    ' On error, don't suspend - let normal processing continue
    DetectAndSuspendBulkOperation = False
End Function

' Suspends change tracking and schedules re-registration on idle
Private Sub SuspendAndScheduleResume()
    On Error GoTo ErrorHandler

    ' Defense in depth: a resume is already pending. Re-running here would overwrite moReRegHandler and
    ' orphan it - DetectAndSuspendBulkOperation's guard normally prevents reaching this point.
    If Not moReRegHandler Is Nothing Then Exit Sub

    MarkChangeTrackingSuspended
    DetachChangeTracking    ' idempotent; keeps BootLoader's attachment state in sync for the scheduled resume

    ' Remove the idle handler from MicroStation's idle list (undefined firing order if two coexist).
    ' Instance kept alive for reuse.
    If mbIdleHandlerRegistered And Not moIdleHandler Is Nothing Then
        RemoveEnterIdleEventHandler moIdleHandler
        mbIdleHandlerRegistered = False
    End If

    ElementInProcesse.Clear    ' pending elements would be stale after the bulk operation

    ' Stored as member (not local var) to prevent premature VBA GC before idle fires
    Set moReRegHandler = New ReRegisterIdleHandler
    AddEnterIdleEventHandler moReRegHandler

    mlCallCount = 0
    mdEntryTime = 0
    mlCachedThreshold = 0
    mlCachedInterval = 0

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.SuspendAndScheduleResume"
End Sub

' ========================================
' PRIVATE HELPER FUNCTIONS - Performance Filtering
' ========================================

' True if at least one ARES feature (tagging, calculation, actuator or rendering) is enabled.
Private Function IsAnyFeatureEnabled() As Boolean
    Dim autoPropertiesEnabled As Boolean
    Dim calcEnabled As Boolean
    Dim renderEnabled As Boolean
    Dim actuatorEnabled As Boolean

    autoPropertiesEnabled = CBool(ARESConfig.ARES_AUTO_PROPERTIES.Value)
    calcEnabled = CBool(ARESConfig.ARES_PROPERTY_CALC.Value)
    renderEnabled = CBool(ARESConfig.ARES_TEXT_RENDER.Value)
    actuatorEnabled = CBool(ARESConfig.ARES_ACTUATE_COLOR.Value) Or CBool(ARESConfig.ARES_ACTUATE_LEVEL.Value)

    IsAnyFeatureEnabled = autoPropertiesEnabled Or calcEnabled Or renderEnabled Or actuatorEnabled
End Function

' True if oElement should be queued for Add/Modify: graphical AND (grouped, OR matches a tag rule now,
' OR already carries an ARES property from an earlier pass). The last check is a deliberately wide
' HasItems("*") match on ANY attached property, not narrowed to what Calc/Render/Actuator actually use -
' do not narrow it without sign-off.
Private Function ShouldQueueElement(ByVal oElement As element) As Boolean
    ShouldQueueElement = False

    If Not oElement.IsGraphical Then Exit Function

    ' The graphic group links a label to its geometry - the cheap pre-filter shared by colour sync,
    ' GroupLength calculation and the renderer's group-driven refresh. Not feature-specific; do not
    ' narrow or remove it when one consumer changes.
    If oElement.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID Then
        ShouldQueueElement = True
        Exit Function
    End If

    If CBool(ARESConfig.ARES_AUTO_PROPERTIES.Value) Then
        If PropertyTagging.ElementMatchesAnyRule(oElement) Then
            ShouldQueueElement = True
            Exit Function
        End If
    End If

    If PropertyCalculation.IsEnabled Or PropertyRendering.IsEnabled Or PropertyActuator.IsEnabled Then
        On Error Resume Next
        oElement.Items.Refresh ARESConstants.ARES_NAME_LIBRARY_TYPE
        ShouldQueueElement = oElement.Items.HasItems(ARESConstants.ARES_NAME_LIBRARY_TYPE, "*")
        If Err.Number <> 0 Then ShouldQueueElement = False
        On Error GoTo 0
    End If
End Function

' True if oElement's linked elements should be queued on deletion: graphical, non-default graphic group,
' and has links. Also returns the linked elements to avoid a second Link.GetLink call.
Private Function ShouldQueueForDeletion(ByVal oElement As element, ByRef linkedEls() As element) As Boolean
    On Error GoTo ErrorHandler

    ShouldQueueForDeletion = False

    If Not oElement.IsGraphical Then Exit Function
    If oElement.GraphicGroup = ARES_DEFAULT_GRAPHIC_GROUP_ID Then Exit Function

    linkedEls = Link.GetLink(oElement)
    If HasElements(linkedEls) Then
        ShouldQueueForDeletion = True
    End If

    Exit Function

ErrorHandler:
    ' On error, don't queue to avoid cascading errors
    ShouldQueueForDeletion = False
End Function

' True if arr is a non-empty, initialized element array (UBound raises on an uninitialized array).
Private Function HasElements(ByRef arr() As element) As Boolean
    On Error Resume Next
    HasElements = False
    If UBound(arr) <> -1 Then
        HasElements = True
    End If
    On Error GoTo 0
End Function
