' Class Module: ElementChangeHandler
' Description: Handles element change events for property tagging, property calculation, property actuation
'              and property rendering (capture + deferred idle processing). Colour sync (linked geometry ->
'              linked text) is now a byproduct of property actuation's GroupColor pilot property, not a
'              standalone concept here.
'              Implements IChangeTrackEvents to listen for element additions, modifications, and deletions.
'              Includes automatic detection of bulk operations using frequency-based detection.
'              DIAGNOSTIC SWITCH (ARES_DIRECT_PROCESSING, default False): when True, an element is processed
'              synchronously inline (ProcessDirect) instead of being queued for the next MicroStation idle
'              event. This exists to let Asketyll test, in real conditions, whether a historical ATLAS
'              instability - the reason deferred/idle processing was introduced (git commit fa361c9,
'              17 Nov 2025; before it, ElementChanged called element handling directly) - still applies now
'              that custom-property processing replaced Auto Lengths. It is a TEST TOOL, not a replacement:
'              the deferred path stays the reference behaviour and default. Direct mode carries a real,
'              different risk from deferred mode: ProcessElement can synchronously .Rewrite an element
'              (PropertyRendering/PropertyActuator/PropertyCalculation), and if MicroStation redispatches
'              ElementChanged on that write before the outer call returns, that re-entry lands back on
'              IChangeTrackEvents_ElementChanged itself - a nested COM call, not the top-level dispatch that
'              deferred mode's own self-triggered cascades get (those happen outside any live ElementChanged
'              call, during idle). mbProcessingDirect below exists specifically to bound that: ProcessElement's
'              own MAX_DEPTH cannot see a re-entry through the event entry point, since each re-entry starts
'              fresh at Depth 0 (a skipped reentrant call is logged once per session, mbReentrancyLogged, so
'              a real hit during testing is greppable). mbProcessingDirect also suppresses
'              DetectAndSuspendBulkOperation while active, mirroring IsIdleProcessingActive() for the idle
'              path: a group's own Calculation/Actuator/Rendering cascade across several siblings within one
'              ProcessDirect call must not be miscounted as an external bulk operation and false-suspend
'              change tracking mid-test.
'              WHAT HAPPENS TO A SKIPPED ELEMENT (confirmed live 2026-08-25, one real reentrant hit logged,
'              no crash): a reentrant call falls back to the deferred queue instead of being dropped -
'              no durable loss either way. CASE 1, the skipped element is ARES's OWN cascade sibling
'              (reached via Link.GetLink / ApplyValueToSibling / PropertyRendering's NoteDirtyGroup+
'              DrainRepaintHop): already had no loss even before the fallback - those are explicit
'              in-call-graph mechanisms independent of ElementInProcesse (true in deferred mode too:
'              Branch 2's own queue-add for a sibling is wiped by IdleEventHandler.CleanupHandler before any
'              fresh pass reaches it; the explicit calls are what actually do the work) - the fallback queue-
'              add here is a harmless redundant no-op for this case. CASE 2, the skipped element is
'              UNRELATED to ARES's own writes (e.g. a concurrent ATLAS edit landing while mbProcessingDirect
'              is True): the fallback is what closes this case - it is picked up on the next idle instead of
'              staying unprocessed until its own next modification.
'              RESULT (Asketyll, live testing, 2026-08-25): the original ATLAS instability this switch was
'              built to test is GONE in direct mode - including under an ATLAS bulk/mass operation, where
'              bulk-suspend and its post-operation resume both fired correctly and performance felt BETTER
'              than deferred, not worse. Root cause unconfirmed either way (Asketyll's own reading: most
'              likely the ATLAS dimension text driving the original slowdown never matches any ARES rule and
'              exits the pipeline immediately, rather than the property-based engines themselves behaving
'              differently from the old Auto Lengths code) - moving to custom-property processing may or may
'              not be what actually fixed it, independently of this switch. Deferred nonetheless stays the
'              shipped default; this switch remains an opt-in diagnostic tool, not a basis for retiring the
'              deferred path.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, ARESConstants, LangManager, StringsInEl,
'               ElementInProcesseClass, ErrorHandlerClass, ReRegisterIdleHandler, PropertyTagging,
'               PropertyCalculation, PropertyActuator, PropertyRendering

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

' Event handler for the beginning of an undo/redo action
' Called before MicroStation processes an undo or redo operation
' Parameters:
'   AfterUndoRedo  - The element state after the undo/redo
'   BeforeUndoRedo - The element state before the undo/redo
'   Action         - The type of change track action
'   IsUndo         - True if this is an undo operation, False for redo
Private Sub IChangeTrackEvents_BeginUndoRedo(ByVal AfterUndoRedo As element, ByVal BeforeUndoRedo As element, ByVal Action As MsdChangeTrackAction, ByVal IsUndo As Boolean)
    On Error GoTo ErrorHandler
    ' Currently not implemented - add code here if undo/redo handling is needed
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.IChangeTrackEvents_BeginUndoRedo"
    Err.Clear
End Sub

' Event handler for when an element is changed (added, modified, or deleted)
' This is the main entry point for tracking element changes in MicroStation
' Parameters:
'   AfterChange   - The element after the change (Nothing for deletions)
'   BeforeChange  - The element before the change (Nothing for additions)
'   Action        - The type of action (Add, Modify, Delete)
'   CantBeUndone  - Whether the change can be undone
Private Sub IChangeTrackEvents_ElementChanged(ByVal AfterChange As element, ByVal BeforeChange As element, ByVal Action As MsdChangeTrackAction, CantBeUndone As Boolean)
    On Error GoTo ErrorHandler
    Dim els() As element
    Dim i As Long

    ' PERFORMANCE: Detect and suspend during bulk operations
    If DetectAndSuspendBulkOperation() Then Exit Sub

    ' GUARD: Config not yet initialized (e.g. seed elements added before first idle fires)
    If Not ARESConfig.IsInitialized Then Exit Sub

    ' PERFORMANCE: Early exit if all ARES features are disabled
    If Not IsAnyFeatureEnabled() Then Exit Sub

    Select Case Action
        ' Handle element addition - queue (or, in direct mode, process inline) the new element
        ' Note: BeforeChange is Nothing for Add actions
        Case msdChangeTrackActionAdd
            If Not AfterChange Is Nothing Then
                If ShouldQueueElement(AfterChange) Then
                    QueueOrProcessDirect AfterChange
                End If
            End If

        ' Handle element modification - queue (or, in direct mode, process inline) the modified element
        Case msdChangeTrackActionModify
            If Not AfterChange Is Nothing Then
                If ShouldQueueElement(AfterChange) Then
                    QueueOrProcessDirect AfterChange
                End If
            End If

        ' Handle element deletion - queue (or, in direct mode, process inline) linked elements for update
        ' Note: AfterChange is Nothing for Delete actions
        Case msdChangeTrackActionDelete
            If Not BeforeChange Is Nothing Then
                If ShouldQueueForDeletion(BeforeChange, els) Then
                    ' Queue/process each linked element. Property Calculation needs no deletion
                    ' hook here (epic 14): deleting a trigger cell already re-queues its former group
                    ' members (via Link.GetLink), so their Depth-0 BEARING pass re-resolves each carried
                    ' calc property against the CURRENT group - the surviving matching cell's text, or ""
                    ' (transition-guarded clear/detach) when none survives. No pending-clear machinery.
                    For i = LBound(els) To UBound(els)
                        QueueOrProcessDirect els(i)
                    Next i
                End If
            End If
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.IChangeTrackEvents_ElementChanged"
    Err.Clear
End Sub

' Event handler for the end of an undo/redo action
' Called after MicroStation completes an undo or redo operation
' Parameters:
'   IsUndo - True if this was an undo operation, False for redo
Private Sub IChangeTrackEvents_FinishUndoRedo(ByVal IsUndo As Boolean)
    On Error GoTo ErrorHandler
    ' Currently not implemented - add code here if post-undo/redo handling is needed
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.IChangeTrackEvents_FinishUndoRedo"
    Err.Clear
End Sub

' Event handler for marking changes (checkpoint in change tracking)
' Called when MicroStation marks a point in the change history
Private Sub IChangeTrackEvents_Mark()
    On Error GoTo ErrorHandler
    ' Currently not implemented - add code here if mark handling is needed
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.IChangeTrackEvents_Mark"
    Err.Clear
End Sub

' ========================================
' PRIVATE METHODS - Direct vs Deferred Dispatch
' ========================================

' Routes one already-filtered element to either the deferred queue (default) or, when
' ARES_DIRECT_PROCESSING is enabled, straight to ProcessDirect. Single call site so Add/Modify/
' Delete-linked all pick up the switch identically - see class header for what direct mode is for.
Private Sub QueueOrProcessDirect(ByVal oElement As element)
    If CBool(ARESConfig.ARES_DIRECT_PROCESSING.Value) Then
        ProcessDirect oElement
    Else
        If ElementInProcesse.Add(oElement) Then EnsureIdleHandlerRegistered
    End If
End Sub

' Processes one element synchronously, inline, from within IChangeTrackEvents_ElementChanged.
' mbProcessingDirect guards against COM reentrancy: ProcessElement can .Rewrite an element
' (PropertyRendering/PropertyActuator/PropertyCalculation), and if that write makes MicroStation
' redispatch ElementChanged before this call returns, the reentry comes back through THIS event
' entry point - a nested call ProcessElement's own MAX_DEPTH cannot see, since a reentrant call
' starts fresh at Depth 0. On a reentrant call, this element is NOT processed synchronously here -
' it falls back to the deferred queue (ElementInProcesse.Add + EnsureIdleHandlerRegistered) instead
' of being dropped, so it is still picked up on the next idle. Safe to call from within the
' reentrant COM frame: unlike DetachChangeTracking (bulk-suspend path), this only touches the idle
' list, not change tracking itself. Idempotent even if the same element is already covered by
' ARES's own explicit cascade (Link.GetLink / ApplyValueToSibling / NoteDirtyGroup+DrainRepaintHop)
' - every engine is compare-before-write/HasItems-guarded, so a redundant idle pass is a cheap no-op,
' not a correctness issue. ElementInProcesse itself is a dedup-by-ID queue, so re-skipping the same
' element multiple times within one reentrant burst cannot duplicate it either.
Private Sub ProcessDirect(ByVal oElement As element)
    On Error GoTo ErrorHandler

    If mbProcessingDirect Then
        ' One-shot log (not status bar - this is a diagnostic switch, Asketyll reads the log, not
        ' the status line) so a real reentrant hit is greppable while testing this switch, without
        ' flooding the log on every subsequent occurrence in the same session.
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

' Ensures the single persistent IdleEventHandler is registered in MicroStation's idle list.
' The instance is created once (lazily) and reused for the lifetime of this ChangeHandler:
' we only add/remove it from the idle list, never destroy and recreate it per batch. This
' avoids ObjPtr address reuse and the stale-handler desync that occurs when a fresh instance
' is created while a previous one is still live and registered.
Private Sub EnsureIdleHandlerRegistered()
    On Error GoTo ErrorHandler

    ' Already in the idle list - nothing to do.
    If mbIdleHandlerRegistered Then Exit Sub

    ' Create the instance only the first time; reuse it on every subsequent batch.
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

' Called by IdleEventHandler (CleanupHandler) after it removes itself from the idle list.
' Clears the "registered" flag so the next batch re-adds the handler. The instance in
' moIdleHandler is kept alive and reused (single persistent handler) - not set to Nothing.
Public Sub ResetIdleHandlerFlag()
    On Error GoTo ErrorHandler

    mbIdleHandlerRegistered = False

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.ResetIdleHandlerFlag"
End Sub

' Returns True if oHandler is the active processing idle handler (the single persistent
' instance in moIdleHandler). Used by IdleEventHandler.CleanupHandler so that a non-active
' instance - the one-shot initial handler created at boot, or an orphan left over after
' ChangeHandler was replaced on file open / license recovery - cannot clear the queue or
' reset the registration flag that belong to the live processing handler.
Public Function IsActiveIdleHandler(ByVal oHandler As IdleEventHandler) As Boolean
    IsActiveIdleHandler = mbIdleHandlerRegistered And (Not moIdleHandler Is Nothing) And (oHandler Is moIdleHandler)
End Function

' ========================================
' PUBLIC METHODS - Element Processing
' ========================================

' Unified handler for element processing (called by IdleEventHandler)
' Parameters:
'   oElement - The element to process
' Processing logic:
'   0. At Depth 0, ahead of both branches: property tagging, then calculation, then rendering
'   1. For text/cell elements: sync colour with the linked geometry
'   2. For geometric elements: find linked text elements and process them recursively
Public Sub ProcessElement(ByVal oElement As element, Optional ByVal Depth As Integer = 0)
    On Error GoTo ErrorHandler

    Const MAX_DEPTH As Integer = 20
    If Depth > MAX_DEPTH Then Exit Sub

    ' === LOCAL VARIABLES ===
    Dim i As Long                   ' Loop counter (Branch 2)
    Dim els() As element            ' Array of linked elements
    Dim linkedElement As element    ' Current linked element being processed

    ' === VALIDATION: Check element is still valid ===
    ' Elements can become invalid if deleted between queueing and processing
    If oElement Is Nothing Then Exit Sub
    If Not IsElementValid(oElement) Then Exit Sub

    ' Auto-attach configured custom properties (PropertyTagging, grammar v2). A plain rule attaches to the
    ' changed element itself; a rule carrying the leading "@" modifier instead fans the properties out to
    ' the OTHER members of the matched element's graphic group and, symmetrically, a grouped element PULLS
    ' the props of any "@" rule matched by another member of its group (membership-convergent, story 15-1:
    ' an element added to an already tagged group is tagged at its own pass, without re-touching the
    ' matching element). Runs BEFORE the graphic-group filter so it also reaches ungrouped elements, and
    ' only at the top level (Depth 0) so linked descendants are not tagged. Attach is idempotent.
    If Depth = 0 Then
        If CBool(ARESConfig.ARES_AUTO_PROPERTIES.Value) Then
            PropertyTagging.ApplyPropertyRules oElement
        End If
        ' Property Calculation runs at Depth 0 too, BEFORE the graphic-group filter below, so the
        ' BEARING pass still reconciles a former sibling whose graphic group may have collapsed to a
        ' single member after its trigger cell was deleted (re-resolving each carried calc property
        ' against the CURRENT group). IsTriggerCell requires a real graphic group internally, so the
        ' calculation path stays naturally inert for ungrouped cells even though the hook runs before the
        ' filter. Ordered AFTER PropertyTagging so both features' Depth-0 work runs.
        If PropertyCalculation.IsEnabled Then
            PropertyCalculation.ProcessElement oElement
        End If
        ' Property Actuator runs right after Calculation and BEFORE Rendering: it is a pure SELF reaction
        ' (reads its own pilot property, already fresh from the Calculation pass above via CellColor/CellLvl
        ' stable-push, and reflects it onto its own Color/Level) - so it needs the value Calculation just
        ' computed, and Rendering (if a token displays that same property) should see the freshly-actuated
        ' state in the SAME pass. No group scan of its own: PropertyCalculation already keeps every attached
        ' member's pilot property current; see epic 16 cahier des charges §3.1/§4.2/§4.3.
        If PropertyActuator.IsEnabled Then
            PropertyActuator.ProcessElement oElement
        End If
        ' Property Rendering closes the Depth-0 sequence: tagging attaches, calculation fills, rendering
        ' writes the text - so an element added to a labelled group is attached, valued AND displayed in
        ' ONE pass, reading values that are fresh from this very pass. It runs at Depth 0 only (Branch 2
        ' recurses at Depth > 0 and would skip this hook, which is exactly why the renderer keeps its own
        ' bounded repaint hop instead of a second call site), and BEFORE the graphic-group filter below,
        ' so an UNGROUPED bound text is still rendered on its own pass - the renderer's binding is the item
        ' attachment, not the graphic group. Note the accepted v1 limit that goes with it: ShouldQueueElement
        ' is untouched (capture hot path), so a bound text only refreshes from a value pushed by ANOTHER
        ' element when it belongs to a REAL graphic group, which both the calc transport and the repaint
        ' hop need. Ungrouped, only direct contact with that text brings it back through here.
        If PropertyRendering.IsEnabled Then
            PropertyRendering.ProcessElement oElement
        End If
    End If

    ' Validate element - skip if it has the default graphic group (not part of ARES workflow)
    If oElement.GraphicGroup = ARES_DEFAULT_GRAPHIC_GROUP_ID Then
        Exit Sub
    End If

    ' A text/cell element (Branch 1's old scope) has nothing left to do here: colour sync used to be handled
    ' by the legacy ONLY_COLOR hook (retired 2026-08-24 - superseded in production by PropertyActuator's
    ' GroupColor pilot property, which repaints a cell's own sub-elements without the header-cascade
    ' corruption ONLY_COLOR was never exposed to - see PropertyActuator.ActuateColor's header comment), and
    ' text/cell elements are not geometric, so Branch 2 below does not apply to them either.
    If IsTextOrCellElement(oElement) Then Exit Sub

    ' ========================================
    ' BRANCH 2: GEOMETRIC ELEMENTS
    ' ========================================
    ' Process geometric elements (line, arc, shape, etc.). Find linked text elements and trigger their
    ' update. Also fires when Property Calculation holds a GroupLength rule: that source needs every OTHER
    ' group member re-queued when the linked geometry itself changes, exactly like Auto Lengths - but must
    ' work even with Auto Lengths OFF. Same for a GroupColor rule (cahier-des-charges-groupcolor-fillmode.md
    ' §3.2): GroupColor's candidate is found by LINK, not by name/level pattern, so - like GroupLength - it
    ' has no "trigger element" a name-pattern scan could detect; re-evaluating every linked sibling inline
    ' here IS its only propagation mechanism.
    Dim bCalcEnabled As Boolean
    bCalcEnabled = PropertyCalculation.IsEnabled
    If Not (bCalcEnabled And (PropertyCalculation.HasGroupLengthRules Or PropertyCalculation.HasGroupColorRules)) Then Exit Sub

    ' Get all elements linked to this geometry
    els = Link.GetLink(oElement)
    If Not HasElements(els) Then Exit Sub

    ' Queue each linked element for processing
    For i = LBound(els) To UBound(els)
        Set linkedElement = els(i)
        ' Validate linked element before processing
        If Not linkedElement Is Nothing Then
            If ElementInProcesse.Add(linkedElement) Then
                ' Recursive call to process the linked text element
                ProcessElement linkedElement, Depth + 1
            End If
            ' The recursive call above is Depth > 0, so it SKIPS the Depth-0-only Property Calculation
            ' hook (line ~249) - and the queue-add just above is wiped wholesale by
            ' IdleEventHandler.CleanupHandler before a fresh Depth-0 pass ever reaches this sibling.
            ' A GroupLength- or GroupColor-fed sibling therefore needs INLINE treatment to recompute
            ' when the linked element itself changes - not just once, on its own attach. Idempotent/
            ' loop-safe via ApplyValueToSibling's compare-before-write (a redundant call here is a
            ' cheap no-op).
            If bCalcEnabled Then PropertyCalculation.ProcessElement linkedElement
        End If
    Next i

    ' Drain the renderer's repaint hop AFTER the inline calc loop above. That loop can reach
    ' ApplyValueToSibling -> PropertyRendering.NoteDirtyGroup, while this element's own Depth-0 render
    ' hook has ALREADY run earlier in this very call. Without the drain, a geometric element sitting
    ' last in the queue leaves its notes behind and the sibling's text stays stale until something
    ' else happens to touch it - and the list must not leak across batches at all.
    If PropertyRendering.IsEnabled Then PropertyRendering.DrainRepaintHop

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.ProcessElement"
    Err.Clear
End Sub

' ========================================
' PRIVATE HELPER FUNCTIONS
' ========================================

' Checks if an element is a text-type or cell element
' These element types can contain text that may have length triggers
' Parameters:
'   element - The element to check
' Returns:
'   True if the element is a Cell, Text, or TextNode element
Private Function IsTextOrCellElement(ByVal element As element) As Boolean
    IsTextOrCellElement = (element.IsCellElement Or element.IsTextElement Or element.IsTextNodeElement)
End Function

' Validates that an element is still valid and accessible
' Elements can become invalid if deleted between queueing and processing
' Parameters:
'   oElement - The element to validate
' Returns:
'   True if the element is valid and can be processed
Private Function IsElementValid(ByVal oElement As element) As Boolean
    On Error GoTo InvalidElement
    Dim nTest As Long

    IsElementValid = False

    ' Try to access a property - will fail if element is invalid/deleted
    nTest = oElement.Type

    ' Check if element is graphical (required for ARES processing)
    If Not oElement.IsGraphical Then Exit Function

    IsElementValid = True
    Exit Function

InvalidElement:
    IsElementValid = False
End Function

' NOTE: IsNumericText lives in StringsInEl (epic 15) and this module no longer calls it - its caller here
' was the Auto Lengths branch, now removed. The note stays because the helper's OWNERSHIP was long
' mis-documented as Auto Lengths': it is the NUMERIC GATE of PropertyRendering's D6/D7/D8 forge protection,
' the first condition of every one of those guards, and PropertyRendering is now its only caller.

' ========================================
' PRIVATE HELPER FUNCTIONS - Bulk Operation Detection
' ========================================

' Detects bulk operations using frequency-based detection
' Returns True if bulk operation detected and handler was suspended
Private Function DetectAndSuspendBulkOperation() As Boolean
    On Error GoTo ErrorHandler
    Dim dCurrentTime As Double
    Dim dElapsedMs As Double

    DetectAndSuspendBulkOperation = False

    ' ARES's own element modifications during idle processing would trigger false bulk detection.
    ' The idle handler sets this flag to suppress suspension during its own processing cycle.
    If IsIdleProcessingActive() Then Exit Function

    ' Mirrors the guard above for direct mode (ARES_DIRECT_PROCESSING): a group's Calculation/
    ' Actuator/Rendering passes can .Rewrite several sibling members from a single ProcessDirect
    ' call, each potentially redispatching ElementChanged. Without this, ARES's own cascade could
    ' cross ARES_Bulk_Threshold and suspend change tracking - a false bulk-suspend caused by ARES
    ' itself, not by an external app, silently invalidating exactly the test this switch exists for.
    If mbProcessingDirect Then Exit Function

    ' Already suspended for a bulk in progress. RemoveChangeTrackEventsHandler is not synchronous
    ' during a tight bulk loop (e.g. a DGN conversion), so in-flight ElementChanged events keep
    ' arriving after suspension. Treat them as "skip" without re-running SuspendAndScheduleResume:
    ' re-running would re-Remove the change handler (unbalanced) and overwrite moReRegHandler,
    ' orphaning the already-scheduled ReRegisterIdleHandler (GC'd while still in the idle list) and
    ' losing the resume -> change tracking stays dead after the bulk.
    If IsChangeTrackingSuspended() Then
        ' Normal case: a resume is genuinely pending for THIS handler -> skip the in-flight event.
        If Not moReRegHandler Is Nothing Then
            DetectAndSuspendBulkOperation = True
            Exit Function
        End If
        ' Self-heal: the suspended flag is set but no resume is pending on this handler. This means
        ' the suspension was orphaned - typically the ChangeHandler instance was replaced on file
        ' open (DGNOpenClose.InitializeChangeHandler) or license recovery, which GC'd the previous
        ' instance together with its scheduled ReRegisterIdleHandler. We are nevertheless receiving
        ' a change event, so a live handler IS registered: clear the stale flag and fall through so
        ' processing is not skipped forever.
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

    ' Calculate elapsed time in milliseconds
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

    ' Defense in depth: a resume is already pending (a ReRegisterIdleHandler is scheduled and
    ' still referenced). Re-running here would overwrite moReRegHandler and orphan the pending
    ' handler. The DetectAndSuspendBulkOperation guard normally prevents reaching this point.
    If Not moReRegHandler Is Nothing Then Exit Sub

    ' Mark as suspended in BootLoader
    MarkChangeTrackingSuspended

    ' Unregister change tracking (idempotent; keeps BootLoader's attachment state in sync so the
    ' scheduled ReRegisterIdleHandler -> ResumeChangeTracking will correctly re-attach afterwards).
    DetachChangeTracking

    ' Remove the idle handler from MicroStation's idle list (prevents coexisting handlers
    ' with undefined firing order). Keep the instance alive for reuse (single persistent handler).
    If mbIdleHandlerRegistered And Not moIdleHandler Is Nothing Then
        RemoveEnterIdleEventHandler moIdleHandler
        mbIdleHandlerRegistered = False
    End If

    ' Clear any pending elements (they'll be stale after bulk operation)
    ElementInProcesse.Clear

    ' Register idle handler to re-enable change tracking
    ' Stored as member (not local var) to prevent premature VBA GC before idle fires
    Set moReRegHandler = New ReRegisterIdleHandler
    AddEnterIdleEventHandler moReRegHandler

    ' Reset detection state for next cycle
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

' Checks if any ARES feature is enabled
' Used for early exit to avoid unnecessary processing
' Returns:
'   True if at least one feature (property tagging, calculation, actuator or rendering) is enabled
Private Function IsAnyFeatureEnabled() As Boolean
    Dim autoPropertiesEnabled As Boolean
    Dim calcEnabled As Boolean
    Dim renderEnabled As Boolean
    Dim actuatorEnabled As Boolean

    ' Config values are stored as Strings, convert to Boolean with CBool
    autoPropertiesEnabled = CBool(ARESConfig.ARES_AUTO_PROPERTIES.Value)
    calcEnabled = CBool(ARESConfig.ARES_PROPERTY_CALC.Value)
    ' Without this term a render-ONLY configuration never reaches ProcessElement and the feature is inert.
    renderEnabled = CBool(ARESConfig.ARES_TEXT_RENDER.Value)
    ' Without this term an actuator-ONLY configuration never reaches ProcessElement and the feature is inert.
    actuatorEnabled = CBool(ARESConfig.ARES_ACTUATE_COLOR.Value) Or CBool(ARESConfig.ARES_ACTUATE_LEVEL.Value)

    IsAnyFeatureEnabled = autoPropertiesEnabled Or calcEnabled Or renderEnabled Or actuatorEnabled
End Function

' Determines if an element should be queued for processing (Add/Modify actions)
' Filters by: graphical + graphic group, OR a matching PropertyTagging rule for ungrouped elements (so
' rule-based auto-attach reaches elements not in a graphic group).
' Note: Link validation is deferred to ProcessElement for performance
' Parameters:
'   oElement - The element to evaluate
' Returns:
'   True if the element should be queued for processing
Private Function ShouldQueueElement(ByVal oElement As element) As Boolean
    ShouldQueueElement = False

    ' Must be graphical
    If Not oElement.IsGraphical Then Exit Function

    ' The graphic group is what links a label to its geometry, so a grouped element is the cheap
    ' pre-filter for colour sync, GroupLength calculation and the renderer's group-driven refresh alike.
    ' The filter is NOT specific to any one feature - do not narrow or remove it when one of them changes.
    If oElement.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID Then
        ShouldQueueElement = True
        Exit Function
    End If

    ' Ungrouped element: queue it if a tag rule matches it NOW, OR if it already carries an ARES
    ' property from an earlier pass (attached by hand through MicroStation's native UI, or by a tag rule
    ' since disabled/edited). Fixes RA7 (epic 14, widened + closed 2026-08-19): an ungrouped element in
    ' that second case used to never be re-queued on Modify, so NONE of Depth-0's work (Tagging,
    ' Calculation, Rendering) ran for it again - not just the narrow self-source case (e.g. Prop[XY]=Coord
    ' not recomputing on a pure move) RA7 originally named. The fallback is O(1) - HasItems(..., "*"),
    ' same cost as the existing CustomPropertyHandler.IsItemAttachedToElement, no per-target loop, no
    ' ItemType scan - and independent of ARES_AUTO_PROPERTIES: Calculation/Rendering act on an
    ' already-attached property whether or not auto-tagging itself is on.
    If CBool(ARESConfig.ARES_AUTO_PROPERTIES.Value) Then
        If PropertyTagging.ElementMatchesAnyRule(oElement) Then
            ShouldQueueElement = True
            Exit Function
        End If
    End If

    ' Gated on having a real DOWNSTREAM CONSUMER (Calculation or Rendering enabled), NOT on the tagging
    ' switch above: a colour-sync-only / tagging-only configuration has nothing that would ever act on
    ' this fallback's result, so paying its COM cost on every ungrouped Modify would be pure waste -
    ' exactly the hot-path cost this file avoids everywhere else. Fail-closed on error, mirroring
    ' CustomPropertyHandler.IsItemAttachedToElement (which wraps the identical Refresh+HasItems pair).
    ' SCOPE, ACCEPTED AS-IS (Asketyll, 2026-08-19): HasItems(..., "*") matches ANY attached ARES property,
    ' not just one Calculation/Rendering would actually act on - including a residual/orphaned one left
    ' by a since-disabled or since-edited rule. On a real working DGN (not a clean test file) this can
    ' re-queue more ungrouped elements than the narrow RA7 case strictly needs, paying a full Depth-0 pass
    ' each Modify for elements whose fallback match will lead nowhere. Idempotent (no data corruption),
    ' but not free. A narrower check (limit to properties Calc/Render actually target) was considered and
    ' explicitly declined - the O(1) simplicity of "*" was judged worth the wider net. This is a DELIBERATE
    ' choice, not an oversight - do not narrow it without Asketyll's sign-off.
    ' PropertyActuator (epic 16) is also a real downstream consumer of an already-attached property: an
    ' UNGROUPED element carrying its pilot property (attached by hand, or by a since-disabled/edited tag
    ' rule) is a legitimate SELF-actuation target - it needs no graphic group at all, unlike Calculation's
    ' Cell* sources. Added to the SAME OR term for the SAME reason Calculation/Rendering are, not a new gate.
    If PropertyCalculation.IsEnabled Or PropertyRendering.IsEnabled Or PropertyActuator.IsEnabled Then
        On Error Resume Next
        oElement.Items.Refresh ARESConstants.ARES_NAME_LIBRARY_TYPE
        ShouldQueueElement = oElement.Items.HasItems(ARESConstants.ARES_NAME_LIBRARY_TYPE, "*")
        If Err.Number <> 0 Then ShouldQueueElement = False
        On Error GoTo 0
    End If
End Function

' Determines if linked elements should be queued when an element is deleted
' Filters by: graphical, graphic group, and existing links
' Also returns the linked elements array to avoid double Link.GetLink call
' Parameters:
'   oElement - The element being deleted
'   linkedEls - Output: array of linked elements (if any)
' Returns:
'   True if linked elements should be queued
Private Function ShouldQueueForDeletion(ByVal oElement As element, ByRef linkedEls() As element) As Boolean
    On Error GoTo ErrorHandler

    ShouldQueueForDeletion = False

    ' Must be graphical
    If Not oElement.IsGraphical Then Exit Function

    ' Ignore elements with default graphic group
    If oElement.GraphicGroup = ARES_DEFAULT_GRAPHIC_GROUP_ID Then Exit Function

    ' Check if element has links (linked elements need to be updated)
    linkedEls = Link.GetLink(oElement)
    ' Check if array is valid and has elements
    If HasElements(linkedEls) Then
        ShouldQueueForDeletion = True
    End If

    Exit Function

ErrorHandler:
    ' On error, don't queue to avoid cascading errors
    ShouldQueueForDeletion = False
End Function

' Helper function to safely check if an element array has elements
' Handles cases where the array may not be properly initialized
' Parameters:
'   arr - The element array to check
' Returns:
'   True if the array is valid and contains at least one element
Private Function HasElements(ByRef arr() As element) As Boolean
    On Error Resume Next
    HasElements = False
    ' UBound returns -1 for empty array, raises error for uninitialized
    If UBound(arr) <> -1 Then
        HasElements = True
    End If
    On Error GoTo 0
End Function

