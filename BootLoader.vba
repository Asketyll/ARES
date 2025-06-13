' Module: BootLoader
' Description: Initializes the VBA project on load.

' Dependencies: DGNOpenClose, ElementChangeHandler

Option Explicit

Public ChangeHandler As ElementChangeHandler
Dim oOpenClose As DGNOpenClose

' Entry point when the project is loaded
Sub OnProjectLoad()
    On Error GoTo ErrorHandler
    
    If ModuleExists("ARES_VAR") Then
        MsgBox "ARES_VAR module is ok !", vbOKOnly
        ARES_VAR.InitMSVars
        ShowStatus "ARES Config with MS Vars Initialized"
    Else
        MsgBox "ARES_VAR module is missing !", vbOKOnly
        GoTo ErrorHandler
    End If
    Set oOpenClose = New DGNOpenClose

    Exit Sub

ErrorHandler:
    MsgBox "Erreur lors du chargement automatique de VBA : " & Err.Description, vbOKOnly
End Sub

'This function check if a module exist
Private Function ModuleExists(moduleName As String) As Boolean
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
End Function
