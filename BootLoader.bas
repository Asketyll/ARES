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
    
    ' Check if the module "ARES_VAR" exists
    If ModuleExists("ARES_VAR") Then
        ' Check if ARES_LANGUAGE is initialized
        If ARES_VAR.ARES_LANGUAGE Is Nothing Then
            ' Initialize MS variables if not already done, LangManager can initialize ARES_VAR if it needs to
            ARES_VAR.InitMSVars
        End If
        ' Notify user about initialization status
        MsgBox GetTranslation("BootUserLangInit") & vbCrLf & GetTranslation("BootMSVarsInit"), vbOKOnly
    Else
        ' Notify user about missing module
        MsgBox GetTranslation("BootUserLangInit") & vbCrLf & GetTranslation("BootMSVarsMissing"), vbOKOnly
        GoTo ErrorHandler
    End If
    
    ' Initialize DGNOpenClose
    Set oOpenClose = New DGNOpenClose
    
    Exit Sub

ErrorHandler:
    ' Notify user about failure and error description
    MsgBox GetTranslation("BootFail") & Err.Description, vbOKOnly
End Sub

' Function to check if a module exists
Private Function ModuleExists(moduleName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim vbProj As Object
    Dim vbComp As Object

    ' Access the active VBA project
    Set vbProj = VBE.ActiveVBProject

    ' Loop through all components in the VBA project
    For Each vbComp In vbProj.VBComponents
        ' Check if the component is a module and if its name matches
        If vbComp.Type = 1 Then
            If vbComp.Name = moduleName Then
                ModuleExists = True
                Exit Function
            End If
        End If
    Next vbComp

    ' If the module is not found
    ModuleExists = False
    Exit Function

ErrorHandler:
    ' Handle any errors silently
    ModuleExists = False
End Function
