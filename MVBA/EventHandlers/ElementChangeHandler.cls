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

' Per-cascade "already processed" set - reset at the top of every Depth=0 ProcessElement entry, seeded with
' the trigger element's own ID, consulted/populated by Branch 2 in BOTH modes. Prevents a same-cascade cycle
' (A->B->A through the graphic-group link graph) from re-running a full pass on an element already handled
' earlier in THIS pass - MAX_DEPTH alone bounds runaway recursion depth but does not stop a short cycle well
' under it from doing redundant work. Mirrors PropertyRendering_StateMachine's ElementIdKey/IsIdDone shape
' (DLongToString, linear scan - a group's member count is always small). Scoped to ONE Depth=0 cascade only:
' does NOT protect against two SEPARATE Depth=0 entries (two native events, or two elements independently
' queued and processed back-to-back in the same IdleEventHandler.ProcessPendingElements batch) reaching each
' other - that is a distinct, harder problem, deliberately not addressed here.
Private mSeenIds() As String
Private mnSeenIds As Long

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
        ' This is the expected, routine path whenever ARES's own Depth-0 pipeline Rewrites the element
        ' being processed (Tagging/Calculation/Rendering all can) - not an anomaly worth logging.
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
' bRunEngines (default True) controls the Tagging/Calculation/Actuator/Rendering block below - it used to be
' hardwired to "Depth = 0", but Branch 2 now needs to give a linked sibling its own full pass, at whatever
' depth it sits, when ARES_Direct_Processing is on (see Branch 2 below). Depth keeps its original, sole
' meaning - the MAX_DEPTH recursion safety net - unrelated to bRunEngines. Every EXISTING call site
' (IdleEventHandler.ProcessPendingElements, ElementChangeHandler.ProcessDirect) omits both optional
' arguments, so this default preserves their behaviour exactly.
Public Sub ProcessElement(ByVal oElement As element, Optional ByVal Depth As Integer = 0, Optional ByVal bRunEngines As Boolean = True)
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

    ' A fresh top-level cascade starts here (Depth=0 is never reached by Branch 2's own recursion any
    ' more - it always passes Depth+1): reset the per-cascade seen-set and seed it with THIS element's own
    ' ID, so a cycle that loops back to the original trigger is caught too, not just sibling-to-sibling
    ' cycles. See mSeenIds's declaration comment above.
    If Depth = 0 Then
        mnSeenIds = 0
        MarkIdSeenThisCascade oElement
    End If

    ' Governed by bRunEngines, not "Depth = 0" any more (see header comment above), BEFORE the graphic-
    ' group filter below (so ungrouped elements are reached too): Tagging attaches -> Calculation fills ->
    ' Actuator paints -> Rendering writes the text. Order matters - each stage reads what the previous one
    ' just wrote, in the SAME pass.
    If bRunEngines Then
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

    Dim bDirectProcessing As Boolean
    bDirectProcessing = CBool(ARESConfig.ARES_DIRECT_PROCESSING.Value)

    For i = LBound(els) To UBound(els)
        Set linkedElement = els(i)
        If Not linkedElement Is Nothing Then
            If Not IsIdSeenThisCascade(linkedElement) Then
                If bDirectProcessing Then
                    ' ARES_Direct_Processing=True: give the sibling its OWN full pass (Tagging+Calc+
                    ' Actuator+Rendering) right now, synchronously, instead of the old partial Calc-only
                    ' nudge plus a guaranteed-redundant later idle pass - it is never added to
                    ' ElementInProcesse any more. Only mark it seen (and only recurse) when the recursive
                    ' call will not immediately no-op via the MAX_DEPTH guard above - "seen" must mean
                    ' "processed", not "attempted", or a diamond-shaped link graph could mark an element
                    ' seen on a too-deep path and never give it its real pass via a shorter one.
                    If Depth + 1 <= MAX_DEPTH Then
                        MarkIdSeenThisCascade linkedElement
                        ProcessElement linkedElement, Depth + 1, True
                    End If
                    ' Accepted asymmetry vs deferred mode (reviewer-10, code review): a sibling reached
                    ' beyond MAX_DEPTH gets NO processing at all here (not even a Calc-only fallback),
                    ' whereas deferred mode's separate inline nudge below is NOT depth-gated and still
                    ' fires. Today's old code had one unconditional Calc-only nudge outside any depth
                    ' check, in both modes; merging that nudge into the now depth-gated full-pass call
                    ' removed it for Direct mode specifically. Never observed in practice (needs a 20+
                    ' link group-chain) - left as-is rather than adding back a partial fallback that would
                    ' re-introduce its own "did this sibling already get a full pass elsewhere" question.
                Else
                    ' ARES_Direct_Processing=False: UNCHANGED, byte-for-byte, from the pre-existing
                    ' behaviour - only the seen-set skip above and the seen-marking below are new, and
                    ' both are no-ops on the common (non-cyclic) case exercised by every existing scenario.
                    ' bRunEngines:=False is NOT the default here on purpose: Depth+1 <> 0 already kept the
                    ' engines from running at this call today (via the old "Depth = 0" gate) - if a future
                    ' edit drops this explicit False (trusting the new default), the engines would start
                    ' running at Depth > 0 in DEFERRED mode too, a real regression. Keep it explicit.
                    If ElementInProcesse.Add(linkedElement) Then
                        If Depth + 1 <= MAX_DEPTH Then
                            MarkIdSeenThisCascade linkedElement
                            ProcessElement linkedElement, Depth + 1, False
                        End If
                    End If
                    ' Depth > 0 skips the engines block, and the queue-add is wiped before a fresh pass
                    ' reaches this sibling - a GroupLength/GroupColor-fed sibling needs this INLINE call.
                    If bCalcEnabled Then PropertyCalculation.ProcessElement linkedElement
                End If
            End If
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

' Has oElement already been fully processed (or is it the Depth=0 trigger itself) in the CURRENT top-level
' cascade? See mSeenIds's declaration comment. Silent on a fault (mirrors PropertyRendering_StateMachine's
' IsIdDone) - an element whose ID cannot be read simply gets no dedup, never a wrongly-skipped pass.
Private Function IsIdSeenThisCascade(ByVal oElement As element) As Boolean
    On Error GoTo ErrorHandler
    Dim sKey As String
    Dim i As Long

    IsIdSeenThisCascade = False
    sKey = DLongToString(oElement.id)
    If Len(sKey) = 0 Then Exit Function

    For i = 0 To mnSeenIds - 1
        If mSeenIds(i) = sKey Then
            IsIdSeenThisCascade = True
            Exit Function
        End If
    Next i
    Exit Function

ErrorHandler:
    IsIdSeenThisCascade = False
End Function

Private Sub MarkIdSeenThisCascade(ByVal oElement As element)
    On Error Resume Next
    Dim sKey As String
    sKey = DLongToString(oElement.id)
    If Len(sKey) = 0 Then Exit Sub

    ReDim Preserve mSeenIds(0 To mnSeenIds)
    mSeenIds(mnSeenIds) = sKey
    mnSeenIds = mnSeenIds + 1
End Sub

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
