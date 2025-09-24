' Module: ConfigurationUI
' Description: Event-driven configuration UI using FileDialogHandler
' License: This project is licensed under the AGPL-3.0.
' Dependencies: FileDialogHandler, WindowsFileDialog, LangManager
Option Explicit

' Enhanced export configuration with event-driven dialog
Public Sub ExportConfigurationUI()
    On Error GoTo ErrorHandler
    
    ' Initialize if needed
    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    
    ' Create dialog handler
    Dim Handler As New FileDialogHandler
    Handler.InitializeExport GetTranslation("ConfigExportTitle"), _
                           WindowsFileDialog.GetDefaultConfigDirectory(), _
                           WindowsFileDialog.GenerateDefaultConfigFileName()
    
    ' Show dialog asynchronously
    WindowsFileDialog.ShowSaveFileDialogAsync GetTranslation("ConfigExportTitle"), _
                                            WindowsFileDialog.GetDefaultConfigDirectory(), _
                                            WindowsFileDialog.GenerateDefaultConfigFileName(), _
                                            Handler
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigurationUI.ExportConfigurationUI"
End Sub

' Enhanced import configuration with event-driven dialog
Public Sub ImportConfigurationUI()
    On Error GoTo ErrorHandler
    
    ' Initialize if needed
    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    
    ' Create dialog handler
    Dim Handler As New FileDialogHandler
    Handler.InitializeImport GetTranslation("ConfigImportTitle"), _
                           WindowsFileDialog.GetDefaultConfigDirectory()
    
    ' Show dialog asynchronously
    WindowsFileDialog.ShowOpenFileDialogAsync GetTranslation("ConfigImportTitle"), _
                                             WindowsFileDialog.GetDefaultConfigDirectory(), _
                                             Handler
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigurationUI.ImportConfigurationUI"
End Sub

' Enhanced configuration summary (unchanged)
Public Sub ShowConfigurationSummaryUI()
    On Error GoTo ErrorHandler
    
    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If
    
    Dim Summary As String
    Summary = ARESConfig.GetConfigSummary()
    
    MsgBox Summary, vbOKOnly + vbInformation, GetTranslation("ConfigSummaryTitle")
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigurationUI.ShowConfigurationSummaryUI"
End Sub