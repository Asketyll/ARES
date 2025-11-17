' Class Module: IdleEventHandler
' Description: This class handles the IEnterIdleEvent to execute code when MicroStation is idle.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass, ARESConfigClass, ElementChangeHandler, ElementInProcesseClass
Option Explicit

Implements IEnterIdleEvent

' === PRIVATE MEMBERS ===
Private mbIsProcessing As Boolean  ' Flag to prevent re-entrance

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
    Dim intProcessedCount As Long
    Dim intErrorCount As Long
    
    ' Set processing flag
    mbIsProcessing = True
    
    ' Get all valid elements (already fresh from MicroStation)
    Set colElements = ElementInProcesse.GetAllElements()
    
    ' Check if collection is empty
    If colElements Is Nothing Or colElements.Count = 0 Then
        CleanupHandler
        Exit Sub
    End If
    
    ' Process each element
    intProcessedCount = 0
    intErrorCount = 0
    
    For Each oElement In colElements
        On Error Resume Next
        
        ' Process the element
        ChangeHandler.ProcessElement oElement
        
        If Err.Number = 0 Then
            intProcessedCount = intProcessedCount + 1
        Else
            intErrorCount = intErrorCount + 1
            ErrorHandler.HandleError "Failed to process element ID: " & DLongToString(oElement.ID) & " - " & Err.Description, _
                                   Err.Number, _
                                   "IdleEventHandler.ProcessPendingElements", _
                                   "WARNING"
            Err.Clear
        End If
    Next oElement
    
    On Error GoTo ErrorHandler
    
    ' Log processing statistics if there were issues
    If intErrorCount > 0 Then
        ErrorHandler.HandleError "Idle processing completed: " & intProcessedCount & " succeeded, " & intErrorCount & " failed", _
                               0, _
                               "IdleEventHandler.ProcessPendingElements", _
                               "INFO"
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
    
    ' Clear the queue
    ElementInProcesse.Clear
    
    ' Unregister this idle handler
    RemoveEnterIdleEventHandler Me
    
    ' Reset the registration flag in ChangeHandler
    ChangeHandler.ResetIdleHandlerFlag
    
    ' Reset processing flag
    mbIsProcessing = False
End Sub

' Clean up when class is destroyed
Private Sub Class_Terminate()
    On Error Resume Next
    mbIsProcessing = False
End Sub