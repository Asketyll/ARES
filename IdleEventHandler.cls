' Class Module: IdleEventHandler
' Description: This class handles the IEnterIdleEvent to execute code when MicroStation is idle.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: None

Option Explicit

Implements IEnterIdleEvent

' Event handler for when MicroStation enters idle state
Private Sub IEnterIdleEvent_EnterIdle(ByVal Reserved As Long)
    On Error GoTo ErrorHandler

    ' Remove the handler to ensure it is only called once
    RemoveEnterIdleEventHandler Me

    ' Execute your code to modify Application.Caption here
    Dim projectName As String
    projectName = Application.FullName
    projectName = Left(projectName, Len(projectName) - 4) & " - [ARES]"
    Application.Caption = projectName
    
    Exit Sub

ErrorHandler:
    ' Handle any errors silently or log them as needed
    Err.Clear
End Sub
