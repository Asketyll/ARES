' Class Module: DGNOpenClose
' Description: Manages events when design files are opened or closed.
' This class module handles the events triggered when a design file is opened or closed in MicroStation.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ElementChangeHandler, LangManager, IdleEventHandler, ElementInProcesseClass, ErrorHandlerClass
Option Explicit

' Declare the application moHooks with events
Dim WithEvents moHooks As Application

' Initialize the class
Private Sub Class_Initialize()
    On Error GoTo ErrorHandler
    ' Set the application moHooks to the current application instance
    Set moHooks = Application
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "DGNOpenClose.Class_Initialize"
End Sub

' Event handler for when a design file is closed
Private Sub moHooks_OnDesignFileClosed(ByVal DesignFileName As String)
    On Error GoTo ErrorHandler
    ' Call a sub or function when a DGN file is closed
    ARESConfig.Initialize
    ErrorHandler.UpdateLogFilePath
    ' Reset bulk suspension state so any stale ReRegisterIdleHandler doesn't double-register on next open
    ResetSuspensionState
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "DGNOpenClose.moHooks_OnDesignFileClosed"
End Sub

' Event handler for when a design file is opened
Private Sub moHooks_OnDesignFileOpened(ByVal DesignFileName As String)
    On Error GoTo ErrorHandler
    ' Call a sub or function when a DGN file is opened
    InitializeChangeHandler
    ElementInProcesse.Reset
    ErrorHandler.UpdateLogFilePath
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "DGNOpenClose.moHooks_OnDesignFileOpened"
End Sub

' Initialize the change handler and add event handlers
Private Sub InitializeChangeHandler()
    On Error GoTo ErrorHandler
    ' Cleanly detach any previous handler, drop leftover bulk-suspend state, and attach a fresh
    ' handler in one consistent step. This is the single source of truth for "a new file is open":
    ' it cannot leave change tracking detached-but-flagged-resumed (the deadlock that killed auto
    ' length updates after a DGN conversion that closed the file mid-bulk-suspend).
    ReinitChangeTrackingForNewFile
    ' Show a status message indicating that the event handlers have been initialized
    If LangManager.IsInit Then
        ShowStatus GetTranslation("DGNOpenCloseEventsInitialized")
    Else
        ShowStatus "Track events element initialized."
    End If
    Exit Sub

ErrorHandler:
    ' Show an error message if the change handler initialization fails
    If LangManager.IsInit Then
        ShowStatus GetTranslation("DGNOpenCloseInitError") & Err.Description
    Else
        ShowStatus "Error initializing DGN Open/Close events: " & Err.Description
    End If
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "DGNOpenClose.InitializeChangeHandler"
End Sub