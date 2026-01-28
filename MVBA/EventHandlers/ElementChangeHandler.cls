' Class Module: ElementChangeHandler
' Description: Handles element change events for automatic length and color updates.
'              Implements IChangeTrackEvents to listen for element additions, modifications, and deletions.
'              Includes automatic detection of bulk operations to suspend tracking for performance.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, AutoLengths, ARESConstants, LangManager, StringsInEl,
'               ElementInProcesseClass, ErrorHandlerClass, ReRegisterIdleHandler

Option Explicit
Implements IChangeTrackEvents

' === PRIVATE MEMBERS ===
Private moIdleHandler As IdleEventHandler       ' Reference to the idle event handler for deferred processing
Private mbIdleHandlerRegistered As Boolean      ' Flag to prevent multiple idle handler registrations

' === BULK OPERATION DETECTION ===
' These members track consecutive empty CommandName events to detect bulk operations
' like GCS reprojection that don't set a CommandName
Private mlEmptyCommandCount As Long             ' Counter for consecutive empty command names
Private Const BULK_THRESHOLD As Long = 100      ' Number of empty commands to trigger bulk mode

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
                    ' Queue each linked element for processing
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

' Ensures that an IdleEventHandler is registered for deferred processing
' Uses a flag to prevent multiple registrations within the same batch of changes
' The idle handler will process all queued elements when MicroStation becomes idle
Private Sub EnsureIdleHandlerRegistered()
    On Error GoTo ErrorHandler

    ' Only register if not already registered (prevents duplicate handlers)
    If Not mbIdleHandlerRegistered Then
        Set moIdleHandler = New IdleEventHandler
        AddEnterIdleEventHandler moIdleHandler
        mbIdleHandlerRegistered = True
    End If

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.EnsureIdleHandlerRegistered"
End Sub

' Called by IdleEventHandler to reset the registration flag after processing
' This allows a new idle handler to be registered for the next batch of changes
Public Sub ResetIdleHandlerFlag()
    On Error GoTo ErrorHandler

    mbIdleHandlerRegistered = False
    Set moIdleHandler = Nothing

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ElementChangeHandler.ResetIdleHandlerFlag"
End Sub

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
Public Sub ProcessElement(ByVal oElement As element)
    On Error GoTo ErrorHandler

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
            ' Extract all text strings from the element
            txts = StringsInEl.GetSetTextsInEl(oElement)
            If Not IsArray(txts) Then Exit Sub
            On Error Resume Next
            If UBound(txts) < LBound(txts) Then Exit Sub
            On Error GoTo ErrorHandler

            ' Load trigger patterns from configuration (e.g., "(Xx_m)" where Xx_ is replaced by length)
            Triggers = Split(ARESConfig.ARES_LENGTH_TRIGGER.Value, ARES_VAR_DELIMITER)

            ' RECALCULATE MODE: Clean up existing numeric values between trigger parts
            ' This allows updating lengths when the linked geometry changes
            For i = LBound(txts) To UBound(txts)
                For j = LBound(Triggers) To UBound(Triggers)
                    ' Split trigger by the ID placeholder (e.g., "(" and "_m)" for "(Xx_m)")
                    SplitedTrigger = Split(Triggers(j), ARESConfig.ARES_LENGTH_TRIGGER_ID.Value)
                    If UBound(SplitedTrigger) = 1 Then
                        ' Find positions of trigger parts in the text
                        FirstPartIndex = InStr(1, txts(i), SplitedTrigger(0))
                        SecondPartIndex = InStr(1, txts(i), SplitedTrigger(1))
                        If FirstPartIndex > 0 And SecondPartIndex > FirstPartIndex Then
                            ' Extract the text between trigger parts (should be the length value)
                            MiddleText = Mid(txts(i), FirstPartIndex + Len(SplitedTrigger(0)), SecondPartIndex - FirstPartIndex - Len(SplitedTrigger(0)))
                            ' If it's numeric, remove it to allow recalculation
                            If IsNumericText(MiddleText) Then
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
                        If InStr(1, txts(i), StringsInEl.RemovePattern(Triggers(j), ARESConfig.ARES_LENGTH_TRIGGER_ID.Value)) > 0 Then
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
        ' Find linked text elements and trigger their update
        If Not UPDATE_LENGTH Then Exit Sub

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
                    ProcessElement linkedElement
                End If
            End If
        Next i
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
    Dim lTest As Long

    IsElementValid = False

    ' Try to access a property - will fail if element is invalid/deleted
    lTest = oElement.Type

    ' Check if element is graphical (required for ARES processing)
    If Not oElement.IsGraphical Then Exit Function

    IsElementValid = True
    Exit Function

InvalidElement:
    IsElementValid = False
End Function

' Validates if a string contains only numeric characters
' Used to identify length values between trigger patterns
' Allowed characters: digits (0-9), spaces, commas, and decimal points
' Parameters:
'   text - The string to validate
' Returns:
'   True if the string contains only numeric characters
Private Function IsNumericText(ByVal text As String) As Boolean
    Dim k As Long
    For k = 1 To Len(text)
        If Not (Mid(text, k, 1) Like "[0-9 ,.]" Or Mid(text, k, 1) = " ") Then
            Exit For
        End If
    Next k
    IsNumericText = (k > Len(text))
End Function

' ========================================
' PRIVATE HELPER FUNCTIONS - Bulk Operation Detection
' ========================================

' Detects bulk operations and suspends change tracking if needed
' Returns True if bulk operation detected and handler was suspended
Private Function DetectAndSuspendBulkOperation() As Boolean
    On Error GoTo ErrorHandler
    Dim strCommandName As String

    DetectAndSuspendBulkOperation = False
    strCommandName = CommandState.CommandName

    ' Check for known bulk operation commands (localized names)
    If IsBulkCommand(strCommandName) Then
        SuspendAndScheduleResume
        DetectAndSuspendBulkOperation = True
        Exit Function
    End If

    ' Track consecutive empty command names (reprojection and other bulk ops)
    If Len(strCommandName) = 0 Then
        mlEmptyCommandCount = mlEmptyCommandCount + 1
        If mlEmptyCommandCount >= BULK_THRESHOLD Then
            SuspendAndScheduleResume
            DetectAndSuspendBulkOperation = True
        End If
    Else
        ' Reset counter on non-empty command
        mlEmptyCommandCount = 0
    End If

    Exit Function

ErrorHandler:
    ' On error, don't suspend - let normal processing continue
    DetectAndSuspendBulkOperation = False
End Function

' Checks if command name matches known bulk operations
Private Function IsBulkCommand(strCommandName As String) As Boolean
    IsBulkCommand = False
    If Len(strCommandName) = 0 Then Exit Function

    Select Case UCase(strCommandName)
        ' French
        Case "FUSIONNER RÉFÉRENCES", "FUSIONNER REFERENCES"
            IsBulkCommand = True
        ' English
        Case "MERGE REFERENCES", "REFERENCE MERGE", "MERGE INTO MASTER"
            IsBulkCommand = True
    End Select
End Function

' Suspends change tracking and schedules re-registration on idle
Private Sub SuspendAndScheduleResume()
    On Error GoTo ErrorHandler
    Dim oReRegHandler As ReRegisterIdleHandler

    ' Mark as suspended in BootLoader
    MarkChangeTrackingSuspended

    ' Unregister change tracking
    If Not ChangeHandler Is Nothing Then
        RemoveChangeTrackEventsHandler ChangeHandler
    End If

    ' Clear any pending elements (they'll be stale after bulk operation)
    ElementInProcesse.Clear

    ' Register idle handler to re-enable change tracking
    Set oReRegHandler = New ReRegisterIdleHandler
    AddEnterIdleEventHandler oReRegHandler

    ' Reset counter
    mlEmptyCommandCount = 0

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
'   True if at least one feature (auto lengths or color sync) is enabled
Private Function IsAnyFeatureEnabled() As Boolean
    Dim autoLengthEnabled As Boolean
    Dim onlyColorEnabled As Boolean

    ' Config values are stored as Strings, convert to Boolean with CBool
    ' Auto lengths requires both flags to be True
    autoLengthEnabled = CBool(ARESConfig.ARES_AUTO_LENGTHS.Value) And CBool(ARESConfig.ARES_UPDATE_LENGTHS.Value)
    onlyColorEnabled = CBool(ARESConfig.ARES_ONLY_COLOR.Value)

    IsAnyFeatureEnabled = autoLengthEnabled Or onlyColorEnabled
End Function

' Determines if an element should be queued for processing (Add/Modify actions)
' Filters by: graphical and graphic group only (fast path)
' Note: Link validation is deferred to ProcessElement for performance
' Parameters:
'   oElement - The element to evaluate
' Returns:
'   True if the element should be queued for processing
Private Function ShouldQueueElement(ByVal oElement As element) As Boolean
    ShouldQueueElement = False

    ' Must be graphical
    If Not oElement.IsGraphical Then Exit Function

    ' Ignore elements with default graphic group (not part of ARES workflow)
    If oElement.GraphicGroup = ARES_DEFAULT_GRAPHIC_GROUP_ID Then Exit Function

    ShouldQueueElement = True
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
