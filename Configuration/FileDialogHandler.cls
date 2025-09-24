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
    Else
        ShowStatus GetTranslation("ConfigExportFailed")
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