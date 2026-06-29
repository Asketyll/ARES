' Class Module: IdleEventHandler
' Description: Handles IEnterIdleEvent to execute deferred element processing when MicroStation is idle.
'              This implements a batch processing pattern to avoid blocking the UI during active modeling.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass, ARESConfigClass, ElementChangeHandler, ElementInProcesseClass

Option Explicit

Implements IEnterIdleEvent

' === PRIVATE MEMBERS ===
Private mbIsProcessing As Boolean    ' Flag to prevent re-entrance during element processing
Private mbIsRevalidating As Boolean  ' Flag to prevent re-entrance during license re-validation

' Event handler for when MicroStation enters idle state
Private Sub IEnterIdleEvent_EnterIdle(ByVal Reserved As Long)
    On Error GoTo ErrorHandler

    ' Protection against re-entrance
    If mbIsProcessing Then Exit Sub

    ' Check if this is the initial idle event (project load)
    If Not LangManager.IsInit Then
        HandleInitialIdle
        Exit Sub
    End If

    ' Ensure ARESConfig is initialized even if LangManager was already init.
    ' This MUST run before RecheckLicenseIfDue so the configured re-check interval
    ' (ARES_License_Recheck_Interval) is read instead of falling back to the default.
    If Not ARESConfig.IsInitialized Then
        ARESConfig.Initialize
    End If

    ' Periodic license re-validation (Issue #2 mitigation, Story 1-2).
    ' Throttled by ARES_LICENSE_RECHECK_INTERVAL inside RecheckLicenseIfDue.
    If Not mbIsRevalidating Then
        mbIsRevalidating = True
        BootLoader.RecheckLicenseIfDue
        mbIsRevalidating = False
    End If

    ' Short-circuit: if license is now invalid, OnLicenseInvalidated has already
    ' drained the queue and removed the change-tracking handler. Nothing to process.
    If Not LicenseManager.IsLicenseValid() Then
        CleanupHandler
        Exit Sub
    End If

    ' Check if there are pending elements to process
    If ElementInProcesse.HasElements Then
        ProcessPendingElements
    Else
        ' No elements to process - clean up immediately
        CleanupHandler
    End If

    Exit Sub

ErrorHandler:
    mbIsProcessing = False
    mbIsRevalidating = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "IdleEventHandler.IEnterIdleEvent_EnterIdle"
    ' Ensure cleanup even on error in main entry point
    CleanupHandler
    Err.Clear
End Sub

' Handle the initial idle event after project load (existing behavior)
Private Sub HandleInitialIdle()
    On Error GoTo ErrorHandler

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

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "IdleEventHandler.HandleInitialIdle"
    RemoveEnterIdleEventHandler Me  ' Cleanup even on error
    Err.Clear
End Sub

' Process all pending elements in the queue
Private Sub ProcessPendingElements()
    On Error GoTo ErrorHandler
    
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
    
    Exit Sub

ErrorHandler:
    ' Ensure cleanup even on error
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "IdleEventHandler.ProcessPendingElements"
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
    mbIsRevalidating = False
End Sub

