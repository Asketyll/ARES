' Class Module: DGNOpenClose
' Description: Manages events when design files are opened or closed.
' This class module handles the events triggered when a design file is opened or closed in MicroStation.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ElementChangeHandler, LangManager, IdleEventHandler

Option Explicit

' Declare the application hooks with events
Dim WithEvents hooks As Application

' Initialize the class
Private Sub Class_Initialize()
    On Error GoTo ErrorHandler

    ' Set the application hooks to the current application instance
    Set hooks = Application

    Exit Sub

ErrorHandler:
    ' Handle any errors silently or log them as needed
    Err.Clear
End Sub

' Event handler for when a design file is closed
Private Sub Hooks_OnDesignFileClosed(ByVal DesignFileName As String)
    On Error GoTo ErrorHandler

    ' Call a sub or function when a DGN file is closed
    ' Example: CleanupResources
    ' MsgBox "Hey! I'm an autorun VBA project of ARES!"

    Exit Sub

ErrorHandler:
    ' Handle any errors silently or log them as needed
    Err.Clear
End Sub

' Event handler for when a design file is opened
Private Sub Hooks_OnDesignFileOpened(ByVal DesignFileName As String)
    On Error GoTo ErrorHandler
    
    ' Call a sub or function when a DGN file is opened
    InitializeChangeHandler

    ' Create an instance of the IdleEventHandler
    Dim oIdleEventHandler As New IdleEventHandler

    ' Register the idle event handler
    AddEnterIdleEventHandler oIdleEventHandler
    
    Exit Sub
    
ErrorHandler:
    ' Handle any errors silently or log them as needed
    Err.Clear
End Sub

' Initialize the change handler and add event handlers
Private Sub InitializeChangeHandler()
    On Error GoTo ErrorHandler

    ' Create a new instance of the ElementChangeHandler
    Set ChangeHandler = New ElementChangeHandler

    ' Add event handlers for tracking changes
    AddChangeTrackEventsHandler ChangeHandler

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
End Sub
