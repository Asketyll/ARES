' ClassModule: DGNOpenClose
' Description: Manages events when design files are opened or closed.

' Dependencies: ElementChangeHandler

Option Explicit

' Declare the application hooks with events
Dim WithEvents Hooks As Application

' Initialize the class
Private Sub Class_Initialize()
    Set Hooks = Application
End Sub

' Event handler for when a design file is closed
Private Sub hooks_OnDesignFileClosed(ByVal DesignFileName As String)
    ' Call a sub or function when a DGN file is closed
    ' Example: CleanupResources
    MsgBox "Hey ! i'm a autorun VBA project of ARES !"
End Sub

' Event handler for when a design file is opened
Private Sub hooks_OnDesignFileOpened(ByVal DesignFileName As String)
    ' Call a sub or function when a DGN file is opened
    InitializeChangeHandler
End Sub

' Initialize the change handler and add event handlers
Private Sub InitializeChangeHandler()
    Set ChangeHandler = New ElementChangeHandler
    AddChangeTrackEventsHandler ChangeHandler
    ShowStatus "Hey ! i'm a autorun VBA project of ARES !"
End Sub
