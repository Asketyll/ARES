' Class Module: IdleEventHandler
' Description: This class handles the IEnterIdleEvent to execute code when MicroStation is idle.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass, ARESConfigClass
Option Explicit

Implements IEnterIdleEvent

' Event handler for when MicroStation enters idle state
Private Sub IEnterIdleEvent_EnterIdle(ByVal Reserved As Long)
    On Error GoTo ErrorHandler
    
    ' Remove the handler to ensure it is only called once
    RemoveEnterIdleEventHandler Me

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
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "IdleEventHandler.IEnterIdleEvent_EnterIdle"
    Err.Clear
End Sub