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

    ' Initialize translations
    LangManager.InitializeTranslations
    
    ' Check if ARES_LANGUAGE is initialized
    If ARES_VAR.ARES_LANGUAGE Is Nothing Then
        ' Initialize MS variables if not already done, LangManager can initialize ARES_VAR if it needs to
        ARES_VAR.InitMSVars
    End If
    
    ' Initialize DGNOpenClose
    Set oOpenClose = New DGNOpenClose
    
    Exit Sub

ErrorHandler:
    ' Notify user about failure and error description
    MsgBox GetTranslation("BootFail") & Err.Description, vbOKOnly
End Sub
