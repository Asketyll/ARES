' Module: WindowsFileDialog
' Description: PowerShell-based file dialog for MicroStation VBA (event-driven)
' License: This project is licensed under the AGPL-3.0.
Option Explicit

' Global handler instance
Public DialogHandler As FileDialogHandler

' Show save file dialog (event-driven)
Public Sub ShowSaveFileDialogAsync(ByVal Title As String, _
                                   ByVal InitialDir As String, _
                                   ByVal DefaultFileName As String, _
                                   ByVal Handler As FileDialogHandler)
    On Error GoTo ErrorHandler
    
    ' Store handler reference
    Set DialogHandler = Handler
    
    ' Launch PowerShell dialog asynchronously
    ExecutePowerShellSaveDialog Title, InitialDir, DefaultFileName
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "WindowsFileDialog.ShowSaveFileDialogAsync"
    If Not DialogHandler Is Nothing Then
        DialogHandler.OnFileDialogCancelled
    End If
End Sub

' Show open file dialog (event-driven)
Public Sub ShowOpenFileDialogAsync(ByVal Title As String, _
                                   ByVal InitialDir As String, _
                                   ByVal Handler As FileDialogHandler)
    On Error GoTo ErrorHandler
    
    ' Store handler reference
    Set DialogHandler = Handler
    
    ' Launch PowerShell dialog asynchronously
    ExecutePowerShellOpenDialog Title, InitialDir
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "WindowsFileDialog.ShowOpenFileDialogAsync"
    If Not DialogHandler Is Nothing Then
        DialogHandler.OnFileDialogCancelled
    End If
End Sub

' Execute PowerShell save dialog
Private Sub ExecutePowerShellSaveDialog(ByVal Title As String, ByVal InitialDir As String, ByVal DefaultFileName As String)
    On Error GoTo ErrorHandler
    
    Dim PSScript As String
    Dim TempScript As String
    Dim CallbackScript As String
    
    ' Create unique temporary script file
    TempScript = Environ("TEMP") & "\ares_save_dialog_" & Format(Now, "hhmmssfffff") & ".ps1"
    CallbackScript = Environ("TEMP") & "\ares_callback_" & Format(Now, "hhmmssfffff") & ".vbs"
    
    ' Escape paths for PowerShell
    Dim SafeTitle As String, SafeInitialDir As String, SafeDefaultFileName As String
    SafeTitle = Replace(Title, "'", "''")
    SafeInitialDir = Replace(InitialDir, "\", "\\")
    SafeDefaultFileName = Replace(DefaultFileName, "'", "''")
    
    ' Build PowerShell script
    PSScript = "Add-Type -AssemblyName System.Windows.Forms" & vbCrLf & _
               "$dialog = New-Object System.Windows.Forms.SaveFileDialog" & vbCrLf & _
               "$dialog.Title = '" & SafeTitle & "'" & vbCrLf & _
               "$dialog.Filter = 'ARES Config (*.cfg)|*.cfg|All Files (*.*)|*.*'" & vbCrLf & _
               "$dialog.DefaultExt = 'cfg'" & vbCrLf & _
               "$dialog.InitialDirectory = '" & SafeInitialDir & "'" & vbCrLf & _
               "$dialog.FileName = '" & SafeDefaultFileName & "'" & vbCrLf & _
               "if($dialog.ShowDialog() -eq 'OK') {" & vbCrLf & _
               "    $result = $dialog.FileName" & vbCrLf & _
               "} else {" & vbCrLf & _
               "    $result = ''" & vbCrLf & _
               "}" & vbCrLf & _
               "& cscript.exe //nologo '" & CallbackScript & "' ""$result""" & vbCrLf
    
    ' Create callback VBScript
    Dim CallbackVBS As String
    CallbackVBS = "Dim objExcel" & vbCrLf & _
                  "Set objExcel = CreateObject(""Excel.Application"")" & vbCrLf & _
                  "objExcel.Run ""'" & Application.VBE.ActiveVBProject.FileName & "'!PowerShellSaveCallback"", WScript.Arguments(0)" & vbCrLf & _
                  "objExcel.Quit" & vbCrLf & _
                  "Set objExcel = Nothing"
    
    ' Write scripts to temp files
    WriteTextFile TempScript, PSScript
    WriteTextFile CallbackScript, CallbackVBS
    
    ' Execute PowerShell script asynchronously
    Shell "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & TempScript & """", vbHide
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "WindowsFileDialog.ExecutePowerShellSaveDialog"
    If Not DialogHandler Is Nothing Then
        DialogHandler.OnFileDialogCancelled
    End If
End Sub

' Execute PowerShell open dialog
Private Sub ExecutePowerShellOpenDialog(ByVal Title As String, ByVal InitialDir As String)
    On Error GoTo ErrorHandler
    
    Dim PSScript As String
    Dim TempScript As String
    Dim CallbackScript As String
    
    ' Create unique temporary script files
    TempScript = Environ("TEMP") & "\ares_open_dialog_" & Format(Now, "hhmmssfffff") & ".ps1"
    CallbackScript = Environ("TEMP") & "\ares_callback_" & Format(Now, "hhmmssfffff") & ".vbs"
    
    ' Escape paths for PowerShell
    Dim SafeTitle As String, SafeInitialDir As String
    SafeTitle = Replace(Title, "'", "''")
    SafeInitialDir = Replace(InitialDir, "\", "\\")
    
    ' Build PowerShell script
    PSScript = "Add-Type -AssemblyName System.Windows.Forms" & vbCrLf & _
               "$dialog = New-Object System.Windows.Forms.OpenFileDialog" & vbCrLf & _
               "$dialog.Title = '" & SafeTitle & "'" & vbCrLf & _
               "$dialog.Filter = 'ARES Config (*.cfg)|*.cfg|All Files (*.*)|*.*'" & vbCrLf & _
               "$dialog.DefaultExt = 'cfg'" & vbCrLf & _
               "$dialog.CheckFileExists = $true" & vbCrLf & _
               "$dialog.Multiselect = $false" & vbCrLf & _
               "$dialog.InitialDirectory = '" & SafeInitialDir & "'" & vbCrLf & _
               "if($dialog.ShowDialog() -eq 'OK') {" & vbCrLf & _
               "    $result = $dialog.FileName" & vbCrLf & _
               "} else {" & vbCrLf & _
               "    $result = ''" & vbCrLf & _
               "}" & vbCrLf & _
               "& cscript.exe //nologo '" & CallbackScript & "' ""$result""" & vbCrLf
    
    ' Create callback VBScript
    Dim CallbackVBS As String
    CallbackVBS = "Dim objExcel" & vbCrLf & _
                  "Set objExcel = CreateObject(""Excel.Application"")" & vbCrLf & _
                  "objExcel.Run ""'" & Application.VBE.ActiveVBProject.FileName & "'!PowerShellOpenCallback"", WScript.Arguments(0)" & vbCrLf & _
                  "objExcel.Quit" & vbCrLf & _
                  "Set objExcel = Nothing"
    
    ' Write scripts to temp files
    WriteTextFile TempScript, PSScript
    WriteTextFile CallbackScript, CallbackVBS
    
    ' Execute PowerShell script asynchronously
    Shell "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & TempScript & """", vbHide
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "WindowsFileDialog.ExecutePowerShellOpenDialog"
    If Not DialogHandler Is Nothing Then
        DialogHandler.OnFileDialogCancelled
    End If
End Sub

' Helper to write text file
Private Sub WriteTextFile(ByVal FilePath As String, ByVal Content As String)
    Dim FileNum As Integer
    FileNum = FreeFile
    Open FilePath For Output As #FileNum
    Print #FileNum, Content
    Close #FileNum
End Sub

' Callback functions called by VBScript
Public Sub PowerShellSaveCallback(ByVal FilePath As String)
    On Error Resume Next
    If Not DialogHandler Is Nothing Then
        DialogHandler.OnFileDialogCompleted FilePath
    End If
End Sub

Public Sub PowerShellOpenCallback(ByVal FilePath As String)
    On Error Resume Next
    If Not DialogHandler Is Nothing Then
        DialogHandler.OnFileDialogCompleted FilePath
    End If
End Sub

' Helper functions
Public Function GetDefaultConfigDirectory() As String
    On Error Resume Next
    If Not ActiveDesignFile Is Nothing Then
        GetDefaultConfigDirectory = ActiveDesignFile.Path
    Else
        GetDefaultConfigDirectory = Environ("USERPROFILE") & "\Documents"
    End If
End Function

Public Function GenerateDefaultConfigFileName(Optional ByVal Prefix As String = "ARES_Config") As String
    GenerateDefaultConfigFileName = Prefix & "_" & Format(Now, "yyyymmdd_hhmmss") & ".cfg"
End Function

