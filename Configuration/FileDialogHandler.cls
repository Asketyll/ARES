' Class Module: FileDialogHandler
' Description: Event-driven file dialog handler for configuration import/export
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, LangManager, ErrorHandlerClass
Option Explicit

' Private members to store dialog context
Private mDialogType As String        ' "Export", "Import", "Backup"
Private mTitle As String
Private mInitialDir As String
Private mDefaultFileName As String

' Initialize the dialog handler
Public Sub InitializeExport(ByVal Title As String, ByVal InitialDir As String, ByVal DefaultFileName As String)
    mDialogType = "Export"
    mTitle = Title
    mInitialDir = InitialDir
    mDefaultFileName = DefaultFileName
End Sub

Public Sub InitializeImport(ByVal Title As String, ByVal InitialDir As String)
    mDialogType = "Import"
    mTitle = Title
    mInitialDir = InitialDir
    mDefaultFileName = ""
End Sub

Public Sub InitializeBackup(ByVal Title As String, ByVal InitialDir As String, ByVal DefaultFileName As String)
    mDialogType = "Backup"
    mTitle = Title
    mInitialDir = InitialDir
    mDefaultFileName = DefaultFileName
End Sub

' Event handler called when PowerShell dialog completes successfully
Public Sub OnFileDialogCompleted(ByVal FilePath As String)
    On Error GoTo ErrorHandler
    
    If Len(Trim(FilePath)) = 0 Then
        OnFileDialogCancelled
        Exit Sub
    End If
    
    ' Process based on dialog type
    Select Case mDialogType
        Case "Export"
            ProcessExportFile FilePath
        Case "Import"
            ProcessImportFile FilePath
        Case "Backup"
            ProcessBackupFile FilePath
        Case Else
            ErrorHandler.HandleError "Unknown dialog type: " & mDialogType, 0, "FileDialogHandler.OnFileDialogCompleted"
    End Select
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FileDialogHandler.OnFileDialogCompleted"
End Sub

' Event handler called when dialog is cancelled
Public Sub OnFileDialogCancelled()
    ShowStatus GetTranslation("ConfigOperationCancelled")
End Sub

' Process export file selection
Private Sub ProcessExportFile(ByVal FilePath As String)
    On Error GoTo ErrorHandler
    
    ' Initialize if needed
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If
    
    ' Export configuration
    If ARESConfig.ExportConfig(FilePath) Then
        ShowStatus GetTranslation("ConfigExportSuccess", FilePath)
        
        ' Show success message with option to open location
        Dim OpenLocation As VbMsgBoxResult
        OpenLocation = MsgBox(GetTranslation("ConfigExportSuccess", FilePath) & vbCrLf & vbCrLf & "Open file location?", _
                             vbYesNo + vbInformation, GetTranslation("ConfigExportTitle"))
        
        If OpenLocation = vbYes Then
            OpenFileLocation FilePath
        End If
    Else
        ShowStatus GetTranslation("ConfigExportFailed")
        MsgBox GetTranslation("ConfigExportFailed"), vbCritical + vbOKOnly, GetTranslation("ConfigExportTitle")
    End If
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FileDialogHandler.ProcessExportFile"
    ShowStatus GetTranslation("ConfigExportFailed") & ": " & Err.Description
End Sub

' Process import file selection
Private Sub ProcessImportFile(ByVal FilePath As String)
    On Error GoTo ErrorHandler
    
    ' Initialize if needed
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If
    
    ' Check if file exists
    If Len(Dir(FilePath)) = 0 Then
        MsgBox GetTranslation("ConfigFileNotFound", FilePath), vbCritical + vbOKOnly, GetTranslation("ConfigImportTitle")
        Exit Sub
    End If
    
    ' Ask about overwriting existing settings
    Dim OverwriteChoice As VbMsgBoxResult
    OverwriteChoice = MsgBox(GetTranslation("ConfigOverwritePrompt"), _
                            vbYesNoCancel + vbQuestion, _
                            GetTranslation("ConfigImportOptions"))
    
    If OverwriteChoice = vbCancel Then
        ShowStatus GetTranslation("ConfigOperationCancelled")
        Exit Sub
    End If
    
    ' Import configuration
    If ARESConfig.ImportConfig(FilePath, (OverwriteChoice = vbYes)) Then
        ShowStatus GetTranslation("ConfigImportSuccess", FilePath)
        MsgBox GetTranslation("ConfigImportSuccess", FilePath), vbInformation + vbOKOnly, GetTranslation("ConfigImportTitle")
    Else
        ShowStatus GetTranslation("ConfigImportFailed")
        MsgBox GetTranslation("ConfigImportFailed"), vbCritical + vbOKOnly, GetTranslation("ConfigImportTitle")
    End If
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FileDialogHandler.ProcessImportFile"
    ShowStatus GetTranslation("ConfigImportFailed") & ": " & Err.Description
End Sub

' Process backup file selection
Private Sub ProcessBackupFile(ByVal FilePath As String)
    On Error GoTo ErrorHandler
    
    ' Initialize if needed
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If
    
    ' Backup configuration
    If ARESConfig.ExportConfig(FilePath) Then
        ShowStatus GetTranslation("ConfigBackupSuccess", FilePath)
        
        ' Ask if user wants to open the backup location
        Dim OpenLocation As VbMsgBoxResult
        OpenLocation = MsgBox(GetTranslation("ConfigBackupSuccess", FilePath) & vbCrLf & vbCrLf & "Open backup location?", _
                             vbYesNo + vbInformation, GetTranslation("ConfigBackupTitle"))
        
        If OpenLocation = vbYes Then
            OpenFileLocation FilePath
        End If
    Else
        ShowStatus GetTranslation("ConfigBackupFailed")
        MsgBox GetTranslation("ConfigBackupFailed"), vbCritical + vbOKOnly, GetTranslation("ConfigBackupTitle")
    End If
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FileDialogHandler.ProcessBackupFile"
    ShowStatus GetTranslation("ConfigBackupFailed") & ": " & Err.Description
End Sub

' Helper function to open file location in Windows Explorer
Private Sub OpenFileLocation(ByVal FilePath As String)
    On Error Resume Next
    Shell "explorer.exe /select,""" & FilePath & """", vbNormalFocus
End Sub