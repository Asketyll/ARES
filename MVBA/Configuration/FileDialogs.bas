' Module: FileDialogs
' Description: Consolidated PowerShell-based file dialogs for ARES configuration
' Replaces: WindowsFileDialog.bas, FileDialogHandler.cls, ConfigurationUI.bas
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, LangManager, ErrorHandlerClass
Option Explicit

' === PUBLIC INTERFACE FOR CONFIGURATION MANAGEMENT ===

' Export configuration with file dialog
Public Sub ExportConfigurationUI()
    On Error GoTo ErrorHandler
    
    ' Initialize if needed
    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If
    
    ' Show save dialog
    Dim filePath As String
    filePath = ShowSaveFileDialog(GetTranslation("ConfigExportTitle"), _
                                 GetDefaultConfigDirectory(), _
                                 GenerateDefaultConfigFileName())
    
    If Len(filePath) > 0 Then
        ' Export configuration
        If ARESConfig.ExportConfig(filePath) Then
            ShowStatus GetTranslation("ConfigExportSuccess", filePath)
        Else
            ShowStatus GetTranslation("ConfigExportFailed")
        End If
    Else
        ShowStatus GetTranslation("ConfigOperationCancelled")
    End If
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FileDialogs.ExportConfigurationUI"
    ShowStatus GetTranslation("ConfigExportFailed") & ": " & Err.Description
End Sub

' Import configuration with file dialog
Public Sub ImportConfigurationUI()
    On Error GoTo ErrorHandler
    
    ' Initialize if needed
    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If
    
    ' Show open dialog
    Dim filePath As String
    filePath = ShowOpenFileDialog(GetTranslation("ConfigImportTitle"), _
                                 GetDefaultConfigDirectory())
    
    If Len(filePath) > 0 Then
        ' Check if file exists
        If Len(Dir(filePath)) = 0 Then
            MsgBox GetTranslation("ConfigFileNotFound", filePath), vbCritical + vbOKOnly, GetTranslation("ConfigImportTitle")
            Exit Sub
        End If
        
        ' Ask about overwriting existing settings
        Dim overwriteChoice As VbMsgBoxResult
        overwriteChoice = MsgBox(GetTranslation("ConfigOverwritePrompt"), _
                                vbYesNoCancel + vbQuestion, _
                                GetTranslation("ConfigImportOptions"))
        
        If overwriteChoice = vbCancel Then
            ShowStatus GetTranslation("ConfigOperationCancelled")
            Exit Sub
        End If
        
        ' Import configuration
        If ARESConfig.ImportConfig(filePath, (overwriteChoice = vbYes)) Then
            ShowStatus GetTranslation("ConfigImportSuccess", filePath)
            MsgBox GetTranslation("ConfigImportSuccess", filePath), vbInformation + vbOKOnly, GetTranslation("ConfigImportTitle")
        Else
            ShowStatus GetTranslation("ConfigImportFailed")
            MsgBox GetTranslation("ConfigImportFailed"), vbCritical + vbOKOnly, GetTranslation("ConfigImportTitle")
        End If
    Else
        ShowStatus GetTranslation("ConfigOperationCancelled")
    End If
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FileDialogs.ImportConfigurationUI"
    ShowStatus GetTranslation("ConfigImportFailed") & ": " & Err.Description
End Sub

' Show configuration summary
Public Sub ShowConfigurationSummaryUI()
    On Error GoTo ErrorHandler
    
    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If
    
    Dim summary As String
    summary = ARESConfig.GetConfigSummary()
    
    MsgBox summary, vbOKOnly + vbInformation, GetTranslation("ConfigSummaryTitle")
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FileDialogs.ShowConfigurationSummaryUI"
End Sub

' === CORE DIALOG FUNCTIONS ===

' Show save file dialog using PowerShell
Public Function ShowSaveFileDialog(ByVal title As String, _
                                  ByVal initialDir As String, _
                                  ByVal defaultFileName As String) As String
    On Error GoTo ErrorHandler
    
    ShowSaveFileDialog = ""
    
    ' Escape special characters for PowerShell
    Dim safeTitle As String, safeInitialDir As String, safeDefaultFileName As String
    safeTitle = EscapeForPowerShell(title)
    safeInitialDir = EscapeForPowerShell(initialDir)
    safeDefaultFileName = EscapeForPowerShell(defaultFileName)
    
    ' Build PowerShell command
    Dim psCommand As String
    psCommand = "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -Command """ & _
                "Add-Type -AssemblyName System.Windows.Forms; " & _
                "$dialog = New-Object System.Windows.Forms.SaveFileDialog; " & _
                "$dialog.Title = '" & safeTitle & "'; " & _
                "$dialog.Filter = 'ARES Config (*.cfg)|*.cfg|All Files (*.*)|*.*'; " & _
                "$dialog.DefaultExt = 'cfg'; " & _
                "$dialog.InitialDirectory = '" & safeInitialDir & "'; " & _
                "$dialog.FileName = '" & safeDefaultFileName & "'; " & _
                "if($dialog.ShowDialog() -eq 'OK') { Write-Output $dialog.FileName }"""
    
    ' Execute command and get result
    Dim result As String
    result = CleanFilePath(GetCommandOutput(psCommand))
    
    ' Return file path if dialog was not cancelled
    If Len(result) > 0 And InStr(result, "ERROR") = 0 Then
        ShowSaveFileDialog = result
    End If
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FileDialogs.ShowSaveFileDialog"
    ShowSaveFileDialog = ""
End Function

' Show open file dialog using PowerShell
Public Function ShowOpenFileDialog(ByVal title As String, _
                                  ByVal initialDir As String) As String
    On Error GoTo ErrorHandler
    
    ShowOpenFileDialog = ""
    
    ' Escape special characters for PowerShell
    Dim safeTitle As String, safeInitialDir As String
    safeTitle = EscapeForPowerShell(title)
    safeInitialDir = EscapeForPowerShell(initialDir)
    
    ' Build PowerShell command
    Dim psCommand As String
    psCommand = "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -Command """ & _
                "Add-Type -AssemblyName System.Windows.Forms; " & _
                "$dialog = New-Object System.Windows.Forms.OpenFileDialog; " & _
                "$dialog.Title = '" & safeTitle & "'; " & _
                "$dialog.Filter = 'ARES Config (*.cfg)|*.cfg|All Files (*.*)|*.*'; " & _
                "$dialog.CheckFileExists = $true; " & _
                "$dialog.Multiselect = $false; " & _
                "$dialog.InitialDirectory = '" & safeInitialDir & "'; " & _
                "if($dialog.ShowDialog() -eq 'OK') { Write-Output $dialog.FileName }"""
    
    ' Execute command and get result
    Dim result As String
    result = CleanFilePath(GetCommandOutput(psCommand))
    
    ' Return file path if dialog was not cancelled
    If Len(result) > 0 And InStr(result, "ERROR") = 0 Then
        ShowOpenFileDialog = result
    End If
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FileDialogs.ShowOpenFileDialog"
    ShowOpenFileDialog = ""
End Function

' === HELPER FUNCTIONS ===

' Execute command and capture output (using working method from test)
Private Function GetCommandOutput(ByVal command As String) As String
    On Error GoTo ErrorHandler
    
    Dim wshShell As Object
    Dim tempFile As String
    Dim batFile As String
    Dim output As String
    Dim fileNum As Integer
    
    Set wshShell = CreateObject("WScript.Shell")
    
    ' Create unique temp files
    tempFile = Environ("TEMP") & "\ares_output_" & Format(Now, "hhmmssfffff") & ".txt"
    batFile = Environ("TEMP") & "\ares_cmd_" & Format(Now, "hhmmssfffff") & ".bat"
    
    ' Create batch file with command
    fileNum = FreeFile
    Open batFile For Output As #fileNum
    Print #fileNum, "@echo off"
    Print #fileNum, command & " > """ & tempFile & """"
    Close #fileNum
    
    ' Execute batch file
    wshShell.Run """" & batFile & """", 0, True
    
    ' Read output
    If Dir(tempFile) <> "" Then
        fileNum = FreeFile
        Open tempFile For Input As #fileNum
        If Not EOF(fileNum) Then
            output = Input(LOF(fileNum), fileNum)
        End If
        Close #fileNum
    End If
    
    ' Cleanup
    On Error Resume Next
    If Dir(tempFile) <> "" Then Kill tempFile
    If Dir(batFile) <> "" Then Kill batFile
    On Error GoTo 0
    
    GetCommandOutput = output
    Exit Function
    
ErrorHandler:
    GetCommandOutput = "ERROR: " & Err.Description
    
    ' Cleanup on error
    On Error Resume Next
    If Dir(tempFile) <> "" Then Kill tempFile
    If Dir(batFile) <> "" Then Kill batFile
End Function

' Escape strings for PowerShell command line
Private Function EscapeForPowerShell(ByVal text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "'", "''")  ' Escape single quotes
    result = Replace(result, """", """""") ' Escape double quotes
    result = Replace(result, "\", "\\")   ' Escape backslashes
    EscapeForPowerShell = result
End Function

' Get default directory for configuration files
Public Function GetDefaultConfigDirectory() As String
    On Error Resume Next
    If Not ActiveDesignFile Is Nothing Then
        GetDefaultConfigDirectory = ActiveDesignFile.Path
    Else
        GetDefaultConfigDirectory = Environ("USERPROFILE") & "\Documents"
    End If
    
    ' Ensure directory exists
    If Len(Dir(GetDefaultConfigDirectory, vbDirectory)) = 0 Then
        GetDefaultConfigDirectory = Environ("TEMP")
    End If
End Function

' Generate default configuration file name
Public Function GenerateDefaultConfigFileName(Optional ByVal prefix As String = "ARES_Config") As String
    GenerateDefaultConfigFileName = prefix & "_" & Format(Now, "yyyymmdd_hhmmss") & ".cfg"
End Function

' Clean file path from unwanted characters
Private Function CleanFilePath(ByVal filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim result As String
    Dim i As Integer
    
    ' Start with trimmed string
    result = Trim(filePath)
    
    ' Remove common control characters
    result = Replace(result, vbCr, "")      ' Carriage return
    result = Replace(result, vbLf, "")      ' Line feed
    result = Replace(result, vbTab, "")     ' Tab
    result = Replace(result, vbNullChar, "") ' Null character
    
    ' Remove any character with ASCII < 32 (control characters)
    Dim cleanResult As String
    cleanResult = ""
    For i = 1 To Len(result)
        If Asc(Mid(result, i, 1)) >= 32 Then
            cleanResult = cleanResult & Mid(result, i, 1)
        End If
    Next i
    
    ' Final trim
    CleanFilePath = Trim(cleanResult)
    Exit Function
    
ErrorHandler:
    CleanFilePath = Trim(filePath) ' Fallback to simple trim
End Function