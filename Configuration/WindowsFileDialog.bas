' Module: WindowsFileDialog
' Description: PowerShell-based file dialog with corrected MicroStation COM callback
' License: This project is licensed under the AGPL-3.0.
Option Explicit

' Global handler for current dialog and result file
Public DialogHandler As FileDialogHandler
Public CurrentResultFile As String

' Show save file dialog (event-driven with VBScript callback)
Public Sub ShowSaveFileDialogAsync(ByVal Title As String, _
                                   ByVal InitialDir As String, _
                                   ByVal DefaultFileName As String, _
                                   ByVal Handler As FileDialogHandler)
    On Error GoTo ErrorHandler
    
    ' Store handler reference
    Set DialogHandler = Handler
    
    ' Launch PowerShell dialog with VBScript callback
    ExecutePowerShellSaveDialog Title, InitialDir, DefaultFileName
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "WindowsFileDialog.ShowSaveFileDialogAsync"
    If Not DialogHandler Is Nothing Then
        DialogHandler.OnFileDialogCancelled
    End If
End Sub

' Show open file dialog (event-driven with VBScript callback)
Public Sub ShowOpenFileDialogAsync(ByVal Title As String, _
                                   ByVal InitialDir As String, _
                                   ByVal Handler As FileDialogHandler)
    On Error GoTo ErrorHandler
    
    ' Store handler reference
    Set DialogHandler = Handler
    
    ' Launch PowerShell dialog with VBScript callback
    ExecutePowerShellOpenDialog Title, InitialDir
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "WindowsFileDialog.ShowOpenFileDialogAsync"
    If Not DialogHandler Is Nothing Then
        DialogHandler.OnFileDialogCancelled
    End If
End Sub

' Modified ExecutePowerShellSaveDialog
Private Sub ExecutePowerShellSaveDialog(ByVal Title As String, ByVal InitialDir As String, ByVal DefaultFileName As String)
    On Error GoTo ErrorHandler
    
    Dim PSCommand As String
    Dim BatchCommand As String
    Dim BatchFile As String
    Dim VBSFile As String
    
    ' Create unique files
    Dim TimeStamp As String
    TimeStamp = Format(Now, "hhmmssfffff")
    CurrentResultFile = Environ("TEMP") & "\ares_result_" & TimeStamp & ".txt"
    BatchFile = Environ("TEMP") & "\ares_monitor_" & TimeStamp & ".bat"
    VBSFile = Environ("TEMP") & "\ares_callback_" & TimeStamp & ".vbs"
    
    ' Escape paths for PowerShell
    Dim SafeTitle As String, SafeInitialDir As String, SafeDefaultFileName As String
    SafeTitle = Replace(Title, "'", "''")
    SafeInitialDir = Replace(InitialDir, "\", "\\")
    SafeDefaultFileName = Replace(DefaultFileName, "'", "''")
    
    ' Build PowerShell command
    PSCommand = "Add-Type -AssemblyName System.Windows.Forms; " & _
                "$dialog = New-Object System.Windows.Forms.SaveFileDialog; " & _
                "$dialog.Title = '" & SafeTitle & "'; " & _
                "$dialog.Filter = 'ARES Config (*.cfg)|*.cfg|All Files (*.*)|*.*'; " & _
                "$dialog.DefaultExt = 'cfg'; " & _
                "$dialog.InitialDirectory = '" & SafeInitialDir & "'; " & _
                "$dialog.FileName = '" & SafeDefaultFileName & "'; " & _
                "if($dialog.ShowDialog() -eq 'OK') { " & _
                "    $dialog.FileName | Out-File -FilePath '" & Replace(CurrentResultFile, "\", "\\") & "' -Encoding ASCII -NoNewline " & _
                "} else { " & _
                "    'CANCELLED' | Out-File -FilePath '" & Replace(CurrentResultFile, "\", "\\") & "' -Encoding ASCII -NoNewline " & _
                "}; " & _
                "& """ & BatchFile & """"
    
    ' Create VBScript with your corrected version
    Dim VBSContent As String
    VBSContent = "Option Explicit" & vbCrLf & _
                 "Dim oConnector, msApp" & vbCrLf & _
                 "' Use ApplicationObjectConnector to get active MicroStation" & vbCrLf & _
                 "Set oConnector = GetObject(, ""MicroStationDGN.ApplicationObjectConnector"")" & vbCrLf & _
                 "If Err.Number <> 0 Then" & vbCrLf & _
                 "    WScript.Echo ""No active MicroStation ApplicationObjectConnector found""" & vbCrLf & _
                 "    WScript.Quit" & vbCrLf & _
                 "End If" & vbCrLf & _
                 "Set msApp = oConnector.Application" & vbCrLf & _
                 "' Execute VBA function" & vbCrLf & _
                 "msApp.CommandState.StartDefaultCommand" & vbCrLf & _
                 "msApp.CadInputQueue.SendCommand(""macro vba run WindowsFileDialog.PowerShellDialogCallback"")" & vbCrLf & _
                 "If Err.Number <> 0 Then" & vbCrLf & _
                 "    WScript.Echo ""VBA execution failed: "" & Err.Description" & vbCrLf & _
                 "End If" & vbCrLf & _
                 "On Error GoTo 0" & vbCrLf & _
                 "Set msApp = Nothing" & vbCrLf & _
                 "Set oConnector = Nothing"
    
    ' Create batch monitor (without deleting ResultFile)
    BatchCommand = "@echo off" & vbCrLf & _
                   "timeout /t 1 /nobreak > nul" & vbCrLf & _
                   "cscript //nologo """ & VBSFile & """" & vbCrLf & _
                   "timeout /t 2 /nobreak > nul" & vbCrLf & _
                   "del /q """ & VBSFile & """ 2>nul" & vbCrLf & _
                   "del /q """ & BatchFile & """ 2>nul"
    
    ' Write files and execute
    WriteTextFile VBSFile, VBSContent
    WriteTextFile BatchFile, BatchCommand
    Shell "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -Command """ & PSCommand & """", vbHide
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "WindowsFileDialog.ExecutePowerShellSaveDialog"
    If Not DialogHandler Is Nothing Then
        DialogHandler.OnFileDialogCancelled
    End If
End Sub

' Execute PowerShell open dialog with VBScript callback
Private Sub ExecutePowerShellOpenDialog(ByVal Title As String, ByVal InitialDir As String)
    On Error GoTo ErrorHandler
    
    Dim PSCommand As String
    Dim BatchCommand As String
    Dim BatchFile As String
    Dim VBSFile As String
    
    ' Create unique files
    Dim TimeStamp As String
    TimeStamp = Format(Now, "hhmmssfffff")
    CurrentResultFile = Environ("TEMP") & "\ares_result_" & TimeStamp & ".txt"
    BatchFile = Environ("TEMP") & "\ares_monitor_" & TimeStamp & ".bat"
    VBSFile = Environ("TEMP") & "\ares_callback_" & TimeStamp & ".vbs"
    
    ' Escape paths for PowerShell
    Dim SafeTitle As String, SafeInitialDir As String
    SafeTitle = Replace(Title, "'", "''")
    SafeInitialDir = Replace(InitialDir, "\", "\\")
    
    ' Build PowerShell command
    PSCommand = "Add-Type -AssemblyName System.Windows.Forms; " & _
                "$dialog = New-Object System.Windows.Forms.OpenFileDialog; " & _
                "$dialog.Title = '" & SafeTitle & "'; " & _
                "$dialog.Filter = 'ARES Config (*.cfg)|*.cfg|All Files (*.*)|*.*'; " & _
                "$dialog.DefaultExt = 'cfg'; " & _
                "$dialog.CheckFileExists = $true; " & _
                "$dialog.Multiselect = $false; " & _
                "$dialog.InitialDirectory = '" & SafeInitialDir & "'; " & _
                "if($dialog.ShowDialog() -eq 'OK') { " & _
                "    $dialog.FileName | Out-File -FilePath '" & Replace(CurrentResultFile, "\", "\\") & "' -Encoding ASCII -NoNewline " & _
                "} else { " & _
                "    'CANCELLED' | Out-File -FilePath '" & Replace(CurrentResultFile, "\", "\\") & "' -Encoding ASCII -NoNewline " & _
                "}; " & _
                "& """ & BatchFile & """"
    
    ' Create VBScript with your corrected version
    Dim VBSContent As String
    VBSContent = "Option Explicit" & vbCrLf & _
                 "Dim oConnector, msApp" & vbCrLf & _
                 "' Use ApplicationObjectConnector to get active MicroStation" & vbCrLf & _
                 "Set oConnector = GetObject(, ""MicroStationDGN.ApplicationObjectConnector"")" & vbCrLf & _
                 "If Err.Number <> 0 Then" & vbCrLf & _
                 "    WScript.Echo ""No active MicroStation ApplicationObjectConnector found""" & vbCrLf & _
                 "    WScript.Quit" & vbCrLf & _
                 "End If" & vbCrLf & _
                 "Set msApp = oConnector.Application" & vbCrLf & _
                 "' Execute VBA function" & vbCrLf & _
                 "msApp.CommandState.StartDefaultCommand" & vbCrLf & _
                 "msApp.CadInputQueue.SendCommand(""macro vba run WindowsFileDialog.PowerShellDialogCallback"")" & vbCrLf & _
                 "If Err.Number <> 0 Then" & vbCrLf & _
                 "    WScript.Echo ""VBA execution failed: "" & Err.Description" & vbCrLf & _
                 "End If" & vbCrLf & _
                 "On Error GoTo 0" & vbCrLf & _
                 "Set msApp = Nothing" & vbCrLf & _
                 "Set oConnector = Nothing"
    
    ' Create batch monitor (without deleting ResultFile)
    BatchCommand = "@echo off" & vbCrLf & _
                   "timeout /t 1 /nobreak > nul" & vbCrLf & _
                   "cscript //nologo """ & VBSFile & """" & vbCrLf & _
                   "timeout /t 2 /nobreak > nul" & vbCrLf & _
                   "del /q """ & VBSFile & """ 2>nul" & vbCrLf & _
                   "del /q """ & BatchFile & """ 2>nul"
    
    ' Write files and execute
    WriteTextFile VBSFile, VBSContent
    WriteTextFile BatchFile, BatchCommand
    Shell "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -Command """ & PSCommand & """", vbHide
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "WindowsFileDialog.ExecutePowerShellOpenDialog"
    If Not DialogHandler Is Nothing Then
        DialogHandler.OnFileDialogCancelled
    End If
End Sub

' Callback function called by VBScript
Public Sub PowerShellDialogCallback()
    On Error GoTo ErrorHandler
    
    If DialogHandler Is Nothing Then Exit Sub
    
    Dim Result As String
    Dim FileNum As Integer
    
    ' Read result from file
    If Dir(CurrentResultFile) <> "" Then
        FileNum = FreeFile
        Open CurrentResultFile For Input As #FileNum
        If Not EOF(FileNum) Then
            Result = Input(LOF(FileNum), FileNum)
        End If
        Close #FileNum
        
        ' Delete the result file
        Kill CurrentResultFile
    End If
    
    ' Process result
    If Result = "CANCELLED" Or Len(Result) = 0 Then
        DialogHandler.OnFileDialogCancelled
    Else
        DialogHandler.OnFileDialogCompleted Result
    End If
    
    ' Clean up reference
    Set DialogHandler = Nothing
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "WindowsFileDialog.PowerShellDialogCallback"
    If Not DialogHandler Is Nothing Then
        DialogHandler.OnFileDialogCancelled
        Set DialogHandler = Nothing
    End If
    
    ' Clean up file if it exists
    On Error Resume Next
    If Dir(CurrentResultFile) <> "" Then Kill CurrentResultFile
    On Error GoTo 0
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

' Helper to write text file
Private Sub WriteTextFile(ByVal FilePath As String, ByVal Content As String)
    Dim FileNum As Integer
    FileNum = FreeFile
    Open FilePath For Output As #FileNum
    Print #FileNum, Content
    Close #FileNum
End Sub