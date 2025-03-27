' Module: BootLoader
' Description: Initializes the VBA project on load.

' Dependencies: DGNOpenClose, ElementChangeHandler

Option Explicit

Public ChangeHandler As ElementChangeHandler
Dim oOpenClose As DGNOpenClose

' Entry point when the project is loaded
Sub OnProjectLoad()
    On Error GoTo ErrorHandler

    Set oOpenClose = New DGNOpenClose

    Exit Sub

ErrorHandler:
    ShowStatus "Erreur lors du chargement automatique de VBA : " & Err.Description
End Sub
