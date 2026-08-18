' Class Module: ElementChangeHandler
' Description: Handles element change events for automatic length/color updates, property tagging,
'              property calculation and property rendering (capture + deferred idle processing).
'              Implements IChangeTrackEvents to listen for element additions, modifications, and deletions.
'              Includes automatic detection of bulk operations using frequency-based detection.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, AutoLengths, ARESConstants, LangManager, StringsInEl,
'               ElementInProcesseClass, ErrorHandlerClass, ReRegisterIdleHandler, PropertyTagging,
'               PropertyCalculation, PropertyRendering

Option Explicit
Implements IChangeTrackEvents

' === PRIVATE MEMBERS ===
Private moIdleHandler As IdleEventHandler       ' Single persistent idle handler instance - created once, reused; only add/removed from the idle list
Private mbIdleHandlerRegistered As Boolean      ' Flag to prevent multiple idle handler registrations
Private moReRegHandler As ReRegisterIdleHandler ' Member-scoped ref to prevent premature VBA GC after bulk suspend

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
        ' Handle element addition - queue the new element for processing
        ' Note: BeforeChange is Nothing for Add actions
        Case msdChangeTrackActionAdd
            If Not AfterChange Is Nothing Then
                If ShouldQueueElement(AfterChange) Then
                    If ElementInProcesse.Add(AfterChange) Then
                        EnsureIdleHandlerRegistered
                    End If
                End If
            End If

        ' Handle element modification - queue the modified element for processing
        Case msdChangeTrackActionModify
            If Not AfterChange Is Nothing Then
                If ShouldQueueElement(AfterChange) Then
                    If ElementInProcesse.Add(AfterChange) Then
                        EnsureIdleHandlerRegistered
                    End If
                End If
            End If

        ' Handle element deletion - queue linked elements for update
        ' Note: AfterChange is Nothing for Delete actions
        Case msdChangeTrackActionDelete
            If Not BeforeChange Is Nothing Then
                If ShouldQueueForDeletion(BeforeChange, els) Then
                    ' Queue each linked element for processing. Property Calculation needs no deletion
                    ' hook here (epic 14): deleting a trigger cell already re-queues its former group
                    ' members (via Link.GetLink), so their Depth-0 BEARING pass re-resolves each carried
                    ' calc property against the CURRENT group - the surviving matching cell's text, or ""
                    ' (transition-guarded clear/detach) when none survives. No pending-clear machinery.
                    For i = LBound(els) To UBound(els)
                        If ElementInProcesse.Add(els(i)) Then
                            EnsureIdleHandlerRegistered
                        End If
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
' Processes a single element for automatic length updates and/or color synchronization
' Parameters:
'   oElement - The element to process
' Processing logic:
'   1. For text/cell elements: Update lengths if triggers are found, or sync colors
'   2. For geometric elements: Find linked text elements and process them recursively
Public Sub ProcessElement(ByVal oElement As element, Optional ByVal Depth As Integer = 0)
    On Error GoTo ErrorHandler

    Const MAX_DEPTH As Integer = 20
    If Depth > MAX_DEPTH Then Exit Sub

    ' === LOCAL VARIABLES ===
    Dim AUTO_LENGTH As Boolean      ' Flag: automatic length addition enabled
    Dim UPDATE_LENGTH As Boolean    ' Flag: automatic length update enabled
    Dim ONLY_COLOR As Boolean       ' Flag: only color synchronization mode
    Dim TriggerFinded As Boolean    ' Flag: trigger pattern found in text
    Dim txts() As String            ' Array of text strings in the element
    Dim Triggers() As String        ' Array of trigger patterns from config
    Dim i As Long, j As Long        ' Loop counters
    Dim SplitedTrigger() As String  ' Trigger split by trigger ID
    Dim FirstPartIndex As Long      ' Position of first part of trigger
    Dim SecondPartIndex As Long     ' Position of second part of trigger
    Dim MiddleText As String        ' Text between trigger parts (the length value)
    Dim els() As element            ' Array of linked elements
    Dim linkedElement As element    ' Current linked element being processed
    Dim subEl As element            ' Sub-element within a cell
    Dim ELEnum As ElementEnumerator ' Enumerator for cell sub-elements
    Dim fillcolor As Long           ' Saved fill color for closed elements

    ' === VALIDATION: Check element is still valid ===
    ' Elements can become invalid if deleted between queueing and processing
    If oElement Is Nothing Then Exit Sub
    If Not IsElementValid(oElement) Then Exit Sub

    ' Load configuration flags
    AUTO_LENGTH = ARESConfig.ARES_AUTO_LENGTHS.Value
    UPDATE_LENGTH = ARESConfig.ARES_UPDATE_LENGTHS.Value
    TriggerFinded = False
    ONLY_COLOR = ARESConfig.ARES_ONLY_COLOR.Value

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

    ' ========================================
    ' BRANCH 1: TEXT/CELL ELEMENTS
    ' ========================================
    If IsTextOrCellElement(oElement) Then

        ' --- Automatic Length Processing ---
        ' Both AUTO_LENGTH and UPDATE_LENGTH must be True to enable automatic length processing
        ' This AND condition ensures that disabling either option will disable the feature
        If (AUTO_LENGTH And UPDATE_LENGTH) Then
            ' Coexistence with Property Rendering (the two text engines must never fight over the same
            ' string): Auto Lengths skips any element the renderer owns. A NESTED If, never an "And"
            ' chain - VBA evaluates both operands, so an "And" would pay IsRenderBound's COM probe even
            ' with the length feature off, and would also run it before the cheap flags. The ONLY_COLOR
            ' block below stays reachable, so a rendered text in a group still syncs its color.
            If Not PropertyRendering.IsRenderBound(oElement) Then
                ' Extract all text strings from the element
                txts = StringsInEl.GetSetTextsInEl(oElement)
                If txts(0) = "" Then Exit Sub

                ' Load trigger patterns from configuration (e.g., "(Xx_m)" where Xx_ is replaced by length)
                Triggers = Split(ARESConfig.ARES_LENGTH_TRIGGER.Value, ARES_VAR_DELIMITER)

                ' Cache trigger ID once — avoids repeated config access inside nested loop
                Dim sTriggerID As String
                sTriggerID = ARESConfig.ARES_LENGTH_TRIGGER_ID.Value

                ' RECALCULATE MODE: Clean up existing numeric values between trigger parts
                ' This allows updating lengths when the linked geometry changes
                For i = LBound(txts) To UBound(txts)
                    For j = LBound(Triggers) To UBound(Triggers)
                        ' Split trigger by the ID placeholder (e.g., "(" and "_m)" for "(Xx_m)")
                        SplitedTrigger = Split(Triggers(j), sTriggerID)
                        If UBound(SplitedTrigger) = 1 Then
                            ' Find positions of trigger parts in the text
                            FirstPartIndex = InStr(1, txts(i), SplitedTrigger(0))
                            SecondPartIndex = InStr(1, txts(i), SplitedTrigger(1))
                            If FirstPartIndex > 0 And SecondPartIndex > FirstPartIndex Then
                                ' Extract the text between trigger parts (should be the length value)
                                MiddleText = Mid(txts(i), FirstPartIndex + Len(SplitedTrigger(0)), SecondPartIndex - FirstPartIndex - Len(SplitedTrigger(0)))
                                ' If it's numeric, remove it to allow recalculation
                                If StringsInEl.IsNumericText(MiddleText) Then
                                    txts(i) = Left(txts(i), FirstPartIndex + Len(SplitedTrigger(0)) - 1) & Mid(txts(i), SecondPartIndex)
                                    StringsInEl.GetSetTextsInEl oElement, Join(txts, ARES_VAR_DELIMITER)
                                    TriggerFinded = True
                                End If
                            End If
                        End If
                    Next j
                Next i

                ' SEARCH MODE: Look for trigger patterns without existing length values
                If Not TriggerFinded Then
                    For i = LBound(txts) To UBound(txts)
                        For j = LBound(Triggers) To UBound(Triggers)
                            ' Check if text contains the trigger pattern (without the ID placeholder)
                            If InStr(1, txts(i), StringsInEl.RemovePattern(Triggers(j), sTriggerID)) > 0 Then
                                TriggerFinded = True
                                Exit For
                            End If
                        Next j
                        If TriggerFinded Then Exit For
                    Next i
                End If

                ' If a trigger was found, delegate to AutoLengths class for length calculation and update
                If TriggerFinded Then
                    Dim oAutoLengths As AutoLengths
                    Set oAutoLengths = New AutoLengths
                    oAutoLengths.Initialize oElement
                    oAutoLengths.UpdateLengths
                    Set oAutoLengths = Nothing
                    Exit Sub
                End If
            End If
        End If

        ' --- Color-Only Synchronization Mode ---
        ' Synchronizes the color of text/cell elements with their linked geometry
        If ONLY_COLOR Then
            ' Get the linked geometric element
            els = Link.GetLink(oElement)
            If Not HasElements(els) Then Exit Sub
            Set linkedElement = els(LBound(els))

            ' Validate linked element
            If linkedElement Is Nothing Then Exit Sub
            If Not IsElementValid(linkedElement) Then Exit Sub

            ' Determine the source color from the linked element
            ' If the source is a ClosedElement with FillMode = 2 (outlined), use its fill color
            ' Exception: if fill color is black (0) or white (255), use outline color instead
            Dim sourceColor As Long
            If linkedElement.IsClosedElement Then
                If linkedElement.AsClosedElement.FillMode = 2 Then
                    fillcolor = linkedElement.AsClosedElement.fillcolor
                    ' Use outline color if fill is black or white
                    If fillcolor = 0 Or fillcolor = 255 Then
                        sourceColor = linkedElement.Color
                    Else
                        sourceColor = fillcolor
                    End If
                Else
                    sourceColor = linkedElement.Color
                End If
            Else
                sourceColor = linkedElement.Color
            End If

            ' Process cell elements - iterate through all sub-elements
            If oElement.IsCellElement Then
                Set ELEnum = oElement.AsCellElement.GetSubElements
                Do While ELEnum.MoveNext
                    Set subEl = ELEnum.Current
                    ' Only update if colors differ
                    If subEl.Color <> sourceColor Then
                        ' Handle closed elements (shapes, ellipses, etc.) specially to preserve fill color
                        ' ClosedElement interface covers all fillable elements: ShapeElement, EllipseElement, etc.
                        ' FillMode = 2 (msdFillModeOutlined) means the element has separate outline and fill colors
                        If subEl.IsClosedElement Then
                            If subEl.AsClosedElement.FillMode = 2 Then
                                ' Save fill color, update outline color, restore fill color
                                fillcolor = subEl.AsClosedElement.fillcolor
                                subEl.Color = sourceColor
                                subEl.AsClosedElement.fillcolor = fillcolor
                            Else
                                ' No fill or solid fill - just update the color
                                subEl.Color = sourceColor
                            End If
                        Else
                            ' Non-closed elements (lines, text, etc.) - simple color update
                            subEl.Color = sourceColor
                        End If
                        subEl.Rewrite
                    End If
                Loop

            ' Process standalone text elements
            ElseIf oElement.IsTextElement Then
                If oElement.AsTextElement.Color <> sourceColor Then
                    oElement.AsTextElement.Color = sourceColor
                    oElement.Rewrite
                End If

            ' Process text node elements (multi-line text)
            ElseIf oElement.IsTextNodeElement Then
                If oElement.AsTextNodeElement.Color <> sourceColor Then
                    oElement.AsTextNodeElement.Color = sourceColor
                    oElement.Rewrite
                End If
            End If
            Exit Sub
        End If

    ' ========================================
    ' BRANCH 2: GEOMETRIC ELEMENTS
    ' ========================================
    Else
        ' Process geometric elements (line, arc, shape, etc.)
        ' Find linked text elements and trigger their update. Also fires when Property Calculation holds a
        ' GroupLength rule: that source needs every OTHER group member re-queued when the linked geometry
        ' itself changes, exactly like Auto Lengths - but must work even with Auto Lengths OFF.
        Dim bCalcEnabled As Boolean
        bCalcEnabled = PropertyCalculation.IsEnabled
        If Not UPDATE_LENGTH And Not (bCalcEnabled And PropertyCalculation.HasGroupLengthRules) Then Exit Sub

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
                ' IdleEventHandler.CleanupHandler before a fresh Depth-0 pass ever reaches this sibling (that
                ' queuing exists for Auto Lengths, whose own trigger-text work runs INLINE here regardless of
                ' Depth). A GroupLength-fed sibling needs the same inline treatment to recompute when the
                ' geometry itself changes - not just once, on its own attach. Idempotent/loop-safe via
                ' ApplyValueToSibling's compare-before-write (a redundant call here is a cheap no-op).
                If bCalcEnabled Then PropertyCalculation.ProcessElement linkedElement
            End If
        Next i

        ' Drain the renderer's repaint hop AFTER the inline calc loop above. That loop can reach
        ' ApplyValueToSibling -> PropertyRendering.NoteDirtyGroup, while this element's own Depth-0 render
        ' hook has ALREADY run earlier in this very call. Without the drain, a geometric element sitting
        ' last in the queue leaves its notes behind and the sibling's text stays stale until something
        ' else happens to touch it - and the list must not leak across batches at all.
        If PropertyRendering.IsEnabled Then PropertyRendering.DrainRepaintHop
    End If

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

' NOTE: IsNumericText moved to StringsInEl (epic 15) and is called from Branch 1 above. Branch 1 is now its
' only caller - PropertyRendering deliberately does not test its own expansions against the legacy trigger
' (the renderer wins on a token-bearing text; see :296). Keep it in StringsInEl and do not reintroduce a
' local copy: it is the shared definition of what AutoLengths owns.

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
'   True if at least one feature (auto lengths, color sync, property tagging or property calculation) is enabled
Private Function IsAnyFeatureEnabled() As Boolean
    Dim autoLengthEnabled As Boolean
    Dim onlyColorEnabled As Boolean
    Dim autoPropertiesEnabled As Boolean
    Dim calcEnabled As Boolean
    Dim renderEnabled As Boolean

    ' Config values are stored as Strings, convert to Boolean with CBool
    ' Auto lengths requires both flags to be True
    autoLengthEnabled = CBool(ARESConfig.ARES_AUTO_LENGTHS.Value) And CBool(ARESConfig.ARES_UPDATE_LENGTHS.Value)
    onlyColorEnabled = CBool(ARESConfig.ARES_ONLY_COLOR.Value)
    autoPropertiesEnabled = CBool(ARESConfig.ARES_AUTO_PROPERTIES.Value)
    calcEnabled = CBool(ARESConfig.ARES_PROPERTY_CALC.Value)
    ' Without this term a render-ONLY configuration never reaches ProcessElement and the feature is inert.
    renderEnabled = CBool(ARESConfig.ARES_TEXT_RENDER.Value)

    IsAnyFeatureEnabled = autoLengthEnabled Or onlyColorEnabled Or autoPropertiesEnabled Or calcEnabled Or renderEnabled
End Function

' Determines if an element should be queued for processing (Add/Modify actions)
' Filters by: graphical + graphic group (AutoLengths fast path), OR a matching PropertyTagging rule
' for ungrouped elements (so rule-based auto-attach reaches elements not in a graphic group).
' Note: Link validation is deferred to ProcessElement for performance
' Parameters:
'   oElement - The element to evaluate
' Returns:
'   True if the element should be queued for processing
Private Function ShouldQueueElement(ByVal oElement As element) As Boolean
    ShouldQueueElement = False

    ' Must be graphical
    If Not oElement.IsGraphical Then Exit Function

    ' Grouped elements drive the AutoLengths workflow (linked via graphic groups)
    If oElement.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID Then
        ShouldQueueElement = True
        Exit Function
    End If

    ' Ungrouped element: queue it only if PropertyTagging is on and a rule matches it.
    ' Documented limitation (epic 14, RA7): an UNGROUPED element already bearing a calc-target property
    ' but matching NO tag rule is not re-queued on its own Modify, so a self-source calc rule (e.g.
    ' Prop[XY]=Coord) does not auto-recompute on a pure move. Handling it would cost a per-distinct-target
    ' IsItemAttachedToElement on every ungrouped add/modify (a hot-path cost) - declined for parsimony.
    If CBool(ARESConfig.ARES_AUTO_PROPERTIES.Value) Then
        ShouldQueueElement = PropertyTagging.ElementMatchesAnyRule(oElement)
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

