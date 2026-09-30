' Module: FileDialogs
' Description: PowerShell-based file dialogs (save/open) for all ARES modules, and the configuration
'              import/export UI. ShowSaveDialog falls back to the active design file folder when no
'              initialDir is given. Accented characters of the ANSI code page are supported in paths and
'              titles, as long as %TEMP% is ASCII (see RunFileDialog).
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass, ConfigThemes, LangManager, ErrorHandlerClass
Option Explicit

' === PUBLIC INTERFACE FOR CONFIGURATION MANAGEMENT ===

' Export configuration with file dialog
Public Sub ExportConfigurationUI()
    On Error GoTo ErrorHandler

    ' Initialize if needed
    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize

    ' The dialog opens among the themes, so an export is listed in the theme window. A folder that cannot be
    ' created (already logged) leaves the dialog on its usual folder.
    Dim initialDir As String
    If ConfigThemes.EnsureThemeFolder() Then initialDir = ConfigThemes.ThemeFolder()

    ' Show save dialog
    Dim filePath As String
    filePath = ShowSaveDialog(GetTranslation("ConfigExportTitle"), _
                             initialDir, _
                             GenerateDefaultConfigFileName(), _
                             DIALOG_FILTER_CFG, "cfg")

    If Len(filePath) > 0 Then
        ' Export configuration
        If ARESConfig.ExportConfig(filePath) Then
            LangManager.ShowStatusText GetTranslation("ConfigExportSuccess", filePath)
        Else
            LangManager.ShowStatusText GetTranslation("ConfigExportFailed")
        End If
    Else
        LangManager.ShowStatusText GetTranslation("ConfigOperationCancelled")
    End If

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FileDialogs.ExportConfigurationUI"
    LangManager.ShowStatusText GetTranslation("ConfigExportFailed")
End Sub

' Import configuration with file dialog
Public Sub ImportConfigurationUI()
    On Error GoTo ErrorHandler

    ' Initialize if needed
    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize

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
            LangManager.ShowStatusText GetTranslation("ConfigOperationCancelled")
            Exit Sub
        End If

        ' Import configuration
        Dim nUnknown As Long
        Dim sFileVersion As String
        Dim sWarnings As String
        If ARESConfig.ImportConfig(filePath, (overwriteChoice = vbYes), nUnknown, sFileVersion) Then
            ' An import is no theme, and may have changed ARES_Language: reload the texts before anything is shown.
            ARESConfig.ARES_THEME_CURRENT.Value = ""
            LangManager.InitializeTranslations
            ConfigThemes.ApplyLoadedConfiguration
            LangManager.ShowStatusText GetTranslation("ConfigImportSuccess", filePath)
            MsgBox GetTranslation("ConfigImportSuccess", filePath), vbInformation + vbOKOnly, GetTranslation("ConfigImportTitle")
            sWarnings = ConfigThemes.DescribeLoadWarnings(nUnknown, sFileVersion)
            If Len(sWarnings) > 0 Then LangManager.ShowStatusText GetTranslation("ConfigImportWarnings", sWarnings)
        Else
            LangManager.ShowStatusText GetTranslation("ConfigImportFailed")
            MsgBox GetTranslation("ConfigImportFailed"), vbCritical + vbOKOnly, GetTranslation("ConfigImportTitle")
        End If
    Else
        LangManager.ShowStatusText GetTranslation("ConfigOperationCancelled")
    End If

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FileDialogs.ImportConfigurationUI"
    LangManager.ShowStatusText GetTranslation("ConfigImportFailed")
End Sub

' === CORE DIALOG FUNCTIONS ===

' Show a save file dialog using PowerShell. Returns the chosen path, "" on cancel.
' fileFilter  : pipe-delimited Windows Forms filter string (e.g. DIALOG_FILTER_CFG)
' defaultExt  : extension without dot (e.g. "cfg", "xlsx")
' initialDir  : starting folder; when empty, falls back to the active design file's
'               folder (or Documents if no file is open).
Public Function ShowSaveDialog(ByVal title As String, _
                               ByVal initialDir As String, _
                               ByVal defaultFileName As String, _
                               ByVal fileFilter As String, _
                               ByVal defaultExt As String) As String
    On Error GoTo ErrorHandler

    ShowSaveDialog = ""

    If Len(initialDir) = 0 Then initialDir = GetDefaultConfigDirectory()

    ShowSaveDialog = RunFileDialog("SaveFileDialog", "", title, initialDir, defaultFileName, fileFilter, defaultExt)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FileDialogs.ShowSaveDialog"
    ShowSaveDialog = ""
End Function

' Show an open file dialog for a configuration file using PowerShell. Returns the chosen path, "" on cancel.
Public Function ShowOpenFileDialog(ByVal title As String, _
                                  ByVal initialDir As String) As String
    On Error GoTo ErrorHandler

    ShowOpenFileDialog = ""
    ShowOpenFileDialog = RunFileDialog("OpenFileDialog", "$dialog.CheckFileExists = $true; $dialog.Multiselect = $false; ", _
                                       title, initialDir, "", DIALOG_FILTER_CFG, "")
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FileDialogs.ShowOpenFileDialog"
    ShowOpenFileDialog = ""
End Function

' === HELPER FUNCTIONS ===

' Run a WinForms file dialog through PowerShell; returns the chosen path, "" on cancel.
' cmd reads the .bat in the OEM code page, so an accent on its command line reaches PowerShell corrupted, and
' PowerShell's stdout comes back in that code page too. Every caller string therefore goes through an ANSI
' parameter file read with -Encoding Default, and the answer comes back the same way: the .bat itself holds only
' the %TEMP% file paths, which must be ASCII. A character outside the ANSI code page does not survive.
' sDialogClass: SaveFileDialog or OpenFileDialog. sSettings: extra ASCII PowerShell statements on $dialog.
Private Function RunFileDialog(ByVal sDialogClass As String, ByVal sSettings As String, _
                               ByVal sTitle As String, ByVal sInitialDir As String, ByVal sFileName As String, _
                               ByVal sFilter As String, ByVal sDefaultExt As String) As String
    On Error GoTo ErrorHandler

    Dim sBase As String
    Dim sInFile As String
    Dim sResultFile As String
    Dim sBatFile As String
    Dim sResult As String
    Dim fileNum As Integer

    RunFileDialog = ""

    ' Unique temp files (CLng(Timer * 1000) gives milliseconds since midnight)
    sBase = Environ("TEMP") & "\ares_dialog_" & CStr(CLng(Timer * 1000))
    sInFile = sBase & "_in.txt"
    sResultFile = sBase & "_result.txt"
    sBatFile = sBase & ".bat"
    DeleteTempFile sResultFile

    ' One value per line, in the order the script reads $p[0] to $p[4]
    fileNum = FreeFile
    Open sInFile For Output As #fileNum
    Print #fileNum, SingleLine(sTitle)
    Print #fileNum, SingleLine(sInitialDir)
    Print #fileNum, SingleLine(sFileName)
    Print #fileNum, SingleLine(sFilter)
    Print #fileNum, SingleLine(sDefaultExt)
    Close #fileNum
    fileNum = 0

    fileNum = FreeFile
    Open sBatFile For Output As #fileNum
    Print #fileNum, "@echo off"
    Print #fileNum, "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -Command """ & _
                    "$p = @(Get-Content -LiteralPath '" & EscapeForPowerShell(sInFile) & "' -Encoding Default); " & _
                    "Add-Type -AssemblyName System.Windows.Forms; " & _
                    "$dialog = New-Object System.Windows.Forms." & sDialogClass & "; " & _
                    "$dialog.Title = $p[0]; " & _
                    "$dialog.InitialDirectory = $p[1]; " & _
                    "$dialog.FileName = $p[2]; " & _
                    "$dialog.Filter = $p[3]; " & _
                    "$dialog.DefaultExt = $p[4]; " & _
                    sSettings & _
                    "if($dialog.ShowDialog() -eq 'OK') { Set-Content -LiteralPath '" & EscapeForPowerShell(sResultFile) & _
                    "' -Value $dialog.FileName -Encoding Default }"""
    Close #fileNum
    fileNum = 0

    CreateObject("WScript.Shell").Run """" & sBatFile & """", 0, True

    ' No result file = the dialog was cancelled
    If TempFileExists(sResultFile) Then
        fileNum = FreeFile
        Open sResultFile For Input As #fileNum
        If Not EOF(fileNum) Then Line Input #fileNum, sResult
        Close #fileNum
        fileNum = 0
    End If

    DeleteTempFile sInFile
    DeleteTempFile sResultFile
    DeleteTempFile sBatFile

    RunFileDialog = CleanFilePath(sResult)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FileDialogs.RunFileDialog"
    If fileNum > 0 Then Close #fileNum
    DeleteTempFile sInFile
    DeleteTempFile sResultFile
    DeleteTempFile sBatFile
    RunFileDialog = ""
End Function

' Escape strings for a single-quoted PowerShell string on the command line
Private Function EscapeForPowerShell(ByVal text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "'", "''")  ' Escape single quotes
    result = Replace(result, """", """""") ' Escape double quotes
    EscapeForPowerShell = result
End Function

' A parameter-file value must stay on its own line.
Private Function SingleLine(ByVal text As String) As String
    SingleLine = Replace(Replace(text, vbCr, " "), vbLf, " ")
End Function

Private Function TempFileExists(ByVal sPath As String) As Boolean
    On Error Resume Next
    TempFileExists = ((GetAttr(sPath) And vbDirectory) = 0)
    If Err.Number <> 0 Then TempFileExists = False
    Err.Clear
End Function

Private Sub DeleteTempFile(ByVal sPath As String)
    On Error Resume Next
    If Len(sPath) > 0 Then
        If TempFileExists(sPath) Then Kill sPath
    End If
    Err.Clear
End Sub

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
