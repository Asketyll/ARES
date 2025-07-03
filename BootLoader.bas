' Module: BootLoader
' Description: Initializes the VBA project on load.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: DGNOpenClose, ElementChangeHandler, LangManager

Option Explicit

Public ChangeHandler As ElementChangeHandler
Dim oOpenClose As DGNOpenClose

' Entry point when the project is loaded
Sub OnProjectLoad()
    On Error GoTo ErrorHandler

    ' Initialize DGNOpenClose
    Set oOpenClose = New DGNOpenClose

    ' Create an instance of the IdleEventHandler
    Dim oIdleEventHandler As New IdleEventHandler

    ' Register the idle event handler
    AddEnterIdleEventHandler oIdleEventHandler
    
    Exit Sub

ErrorHandler:
    ' Notify user about failure and error description
    If LangManager.IsInit Then
        MsgBox GetTranslation("BootFail") & Err.Description, vbOKOnly
    Else
        MsgBox "Error in automatic loading of VBA." & Err.Description, vbOKOnly
    End If
End Sub
