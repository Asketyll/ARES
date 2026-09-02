' Class Module: IdleEventHandler
' Description: Handles IEnterIdleEvent to execute deferred element processing when MicroStation is idle.
'              This implements a batch processing pattern to avoid blocking the UI during active modeling.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass, ARESConfigClass, ElementChangeHandler, ElementInProcesseClass,
'               CallStackClass

Option Explicit

Implements IEnterIdleEvent

' === PRIVATE MEMBERS ===
Private mbIsProcessing As Boolean    ' Flag to prevent re-entrance during element processing
Private mbIsInitialHandler As Boolean ' True ONLY on the boot one-shot instance (see MarkAsInitial)

' Flags this instance as the boot one-shot - the only one that must run HandleInitialIdle.
' Set by BootLoader before registering, and cleared the moment it fires.
'
' Dispatching on a per-instance flag rather than on LangManager.IsInit is deliberate: IsInit means
' "translations are loaded", NOT "the boot idle has not run yet". Any status message shown before
' the first idle self-initialises translations - and DGNOpenClose.InitializeChangeHandler shows one
' on every DGN open, which at MicroStation startup lands BETWEEN OnProjectLoad and the first idle.
' The boot handler then took the element-processing branch and silently lost both the "- [ARES]"
' caption and the startup update check.
Public Sub MarkAsInitial()
    mbIsInitialHandler = True
End Sub

' Event handler for when MicroStation enters idle state
Private Sub IEnterIdleEvent_EnterIdle(ByVal Reserved As Long)
    On Error GoTo ErrorHandler

    ' Protection against re-entrance
    If mbIsProcessing Then Exit Sub

    ' The boot one-shot: cleared FIRST - HandleInitialIdle can show a modal (the update dialog),
    ' and MicroStation keeps pumping idle events while it is up.
    If mbIsInitialHandler Then
        mbIsInitialHandler = False
        HandleInitialIdle
        Exit Sub
    End If

    ' Pushed only past the reentrance/initial-idle guards above.
    Dim bStackPushed As Boolean
    CallStack.Push "IdleEventHandler.IEnterIdleEvent_EnterIdle"
    bStackPushed = True

    ' Ensure ARESConfig is initialized before processing.
    If Not ARESConfig.IsInitialized Then
        ARESConfig.Initialize
    End If

    ' Check if there are pending elements to process
    If ElementInProcesse.HasElements Then
        ProcessPendingElements
    Else
        ' No elements to process - clean up immediately
        CleanupHandler
    End If

    CallStack.Pop
    Exit Sub

ErrorHandler:
    mbIsProcessing = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "IdleEventHandler.IEnterIdleEvent_EnterIdle"
    If bStackPushed Then CallStack.Pop
    ' Ensure cleanup even on error in main entry point
    CleanupHandler
    Err.Clear
End Sub

' Handle the initial idle event after project load (existing behavior)
Private Sub HandleInitialIdle()
    On Error GoTo ErrorHandler
    CallStack.Push "IdleEventHandler.HandleInitialIdle"

    Dim projectName As String
    projectName = Application.FullName
    projectName = Left(projectName, Len(projectName) - 4) & " - [ARES]"
    Application.Caption = projectName

    ' Initialize translations
    LangManager.InitializeTranslations
    
    ' Check if ARES_LANGUAGE is initialized
    If Not ARESConfig.IsInitialized Then
        ' Initialize MS variables if not already done, LangManager can initialize ARESConfig if it needs to
        ARESConfig.Initialize
    End If

    ' Check for available updates — silently skipped on network failure or user preference
    UpdateChecker.CheckForUpdate

    ' Remove the handler after initialization
    RemoveEnterIdleEventHandler Me

    CallStack.Pop
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "IdleEventHandler.HandleInitialIdle"
    CallStack.Pop
    RemoveEnterIdleEventHandler Me  ' Cleanup even on error
    Err.Clear
End Sub

' Process all pending elements in the queue
Private Sub ProcessPendingElements()
    On Error GoTo ErrorHandler
    CallStack.Push "IdleEventHandler.ProcessPendingElements"

    Dim colElements As Collection
    Dim oElement As element
    Dim nProcessedCount As Long
    Dim nErrorCount As Long

    ' Set processing flag and suppress bulk detection during ARES's own modifications
    mbIsProcessing = True
    SetIdleProcessingActive True

    ' Get all valid elements (already fresh from MicroStation)
    Set colElements = ElementInProcesse.GetAllElements()

    ' Check if collection is empty
    If colElements Is Nothing Or colElements.Count = 0 Then
        CleanupHandler
        CallStack.Pop
        Exit Sub
    End If
    
    ' Process each element
    nProcessedCount = 0
    nErrorCount = 0
    
    For Each oElement In colElements
        On Error Resume Next

        ' Process the element
        ChangeHandler.ProcessElement oElement

        If Err.Number = 0 Then
            nProcessedCount = nProcessedCount + 1
        Else
            nErrorCount = nErrorCount + 1
            ErrorHandler.HandleError "Failed to process element ID: " & DLongToString(oElement.ID) & " - " & Err.Description, _
                                   Err.Number, _
                                   Err.Source, _
                                   "IdleEventHandler.ProcessPendingElements"
            Err.Clear
        End If
    Next oElement
    
    On Error GoTo ErrorHandler
    
    ' Log processing statistics if there were issues
    If nErrorCount > 0 Then
        ErrorHandler.HandleError "Idle processing completed: " & nProcessedCount & " succeeded, " & nErrorCount & " failed", _
                               0, _
                               "", _
                               "IdleEventHandler.ProcessPendingElements"
    End If

    ' Always cleanup after processing
    CleanupHandler

    CallStack.Pop
    Exit Sub

ErrorHandler:
    ' Ensure cleanup even on error
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "IdleEventHandler.ProcessPendingElements"
    CallStack.Pop
    CleanupHandler
End Sub

' Centralized cleanup method
Private Sub CleanupHandler()
    On Error Resume Next

    ' Only the active processing instance may drain the queue and reset the registration
    ' flag. With a single persistent idle handler this check is reliable: the steady-state
    ' processing handler always matches ChangeHandler.moIdleHandler, while a non-active
    ' instance (the one-shot initial handler created at boot, or an orphan left after
    ' ChangeHandler was replaced) must remove only itself without disturbing the queue/flag.
    Dim bIsActive As Boolean
    bIsActive = Not ChangeHandler Is Nothing
    If bIsActive Then bIsActive = ChangeHandler.IsActiveIdleHandler(Me)

    RemoveEnterIdleEventHandler Me

    If bIsActive Then
        ElementInProcesse.Clear
        ChangeHandler.ResetIdleHandlerFlag
    End If

    SetIdleProcessingActive False
    mbIsProcessing = False
End Sub

' Clean up when class is destroyed
Private Sub Class_Terminate()
    On Error Resume Next
    SetIdleProcessingActive False
    mbIsProcessing = False
End Sub

