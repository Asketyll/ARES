' Module: Command
' Description: Liste all command
' License: This project is licensed under the AGPL-3.0.
' Dependencies: AutoLengths, BootLoader, LangManager, ARESConfigClass, ConfigurationUI
Option Explicit

' === AUTO LENGTHS COMMANDS ===

' Sub to call CommandState for manual update length in string
Sub ForceUpdateLength()
    CommandState.StartLocate New AutoLengths
End Sub

' === CONFIGURATION MANAGEMENT COMMANDS ===

' Export current configuration using event-driven UI
Sub ExportARESConfig()
    On Error GoTo ErrorHandler
    ConfigurationUI.ExportConfigurationUI
    Exit Sub
    
ErrorHandler:
    If LangManager.IsInit Then
        ShowStatus GetTranslation("ConfigExportFailed") & ": " & Err.Description
    Else
        ShowStatus "Configuration export failed: " & Err.Description
    End If
End Sub

' Import configuration using event-driven UI
Sub ImportARESConfig()
    On Error GoTo ErrorHandler
    ConfigurationUI.ImportConfigurationUI
    Exit Sub
    
ErrorHandler:
    If LangManager.IsInit Then
        ShowStatus GetTranslation("ConfigImportFailed") & ": " & Err.Description
    Else
        ShowStatus "Configuration import failed: " & Err.Description
    End If
End Sub

' Show current configuration summary
Sub ShowARESConfigSummary()
    On Error GoTo ErrorHandler
    ConfigurationUI.ShowConfigurationSummaryUI
    Exit Sub
    
ErrorHandler:
    If LangManager.IsInit Then
        ShowStatus "Failed to show configuration summary: " & Err.Description
    Else
        ShowStatus "Configuration summary failed: " & Err.Description
    End If
End Sub

' Create configuration backup using event-driven UI
Sub BackupARESConfig()
    On Error GoTo ErrorHandler
    ConfigurationUI.BackupConfigurationUI
    Exit Sub
    
ErrorHandler:
    If LangManager.IsInit Then
        ShowStatus GetTranslation("ConfigBackupFailed") & ": " & Err.Description
    Else
        ShowStatus "Configuration backup failed: " & Err.Description
    End If
End Sub

' Quick backup with automatic filename (no dialog)
Sub QuickBackupARESConfig()
    On Error GoTo ErrorHandler
    
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If
    
    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    
    ' Generate automatic backup path
    Dim BackupPath As String
    If Not ActiveDesignFile Is Nothing Then
        BackupPath = ActiveDesignFile.Path & "\ARES_QuickBackup_" & Format(Now, "yyyymmdd_hhmmss") & ".cfg"
    Else
        BackupPath = Environ("USERPROFILE") & "\Desktop\ARES_QuickBackup_" & Format(Now, "yyyymmdd_hhmmss") & ".cfg"
    End If
    
    If ARESConfig.ExportConfig(BackupPath) Then
        ShowStatus "Quick backup created: " & BackupPath
    Else
        ShowStatus GetTranslation("ConfigBackupFailed")
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowStatus "Quick backup failed: " & Err.Description
End Sub

' === VARIABLE MANAGEMENT COMMANDS ===

' Sub to reset all ARES var in MS
Sub ResetARESVariables()
    On Error GoTo ErrorHandler
    
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If
    
    If ARESConfig.ResetAllConfigVars() Then
        If Not LangManager.IsInit Then LangManager.InitializeTranslations
        ShowStatus GetTranslation("VarResetAllSuccess")
    Else
        ShowStatus GetTranslation("VarResetAllFailed")
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowStatus "Reset variables failed: " & Err.Description
End Sub

' Sub to remove all ARES var in MS
Sub RemoveARESVariables()
    On Error GoTo ErrorHandler
    
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If
    
    If ARESConfig.RemoveAllConfigVars() Then
        If Not LangManager.IsInit Then LangManager.InitializeTranslations
        ShowStatus GetTranslation("VarRemoveSuccess")
    Else
        ShowStatus GetTranslation("VarRemoveError")
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowStatus "Remove variables failed: " & Err.Description
End Sub

' === GUI COMMANDS ===

' Sub to call GUI Options of AutoLengths
Sub EditAutoLengthsOptions()
    On Error GoTo ErrorHandler
    
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If
    
    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    
    Dim frm As New AutoLengths_GUI_Options
    frm.Show vbModeless
    
    Exit Sub
    
ErrorHandler:
    ShowStatus "Failed to open AutoLengths options: " & Err.Description
End Sub

' === LANGUAGE COMMANDS ===

' Sub to set language to English
Sub English()
    On Error GoTo ErrorHandler
    
    If Config.SetVar("ARES_Language", "English") Then
        ShowStatus "ARES_Language set to English, please restart."
    Else
        ShowStatus "Impossible to set ARES_Language, please try manually."
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowStatus "Failed to set English language: " & Err.Description
End Sub

' Sub to set language to French
Sub Français()
    On Error GoTo ErrorHandler
    
    If Config.SetVar("ARES_Language", "Français") Then
        ShowStatus "ARES_Language défini à Français, veuillez redémarrer."
    Else
        ShowStatus "Impossible de définir ARES_Language, veuillez essayer manuellement."
    End If
    
    Exit Sub
    
ErrorHandler:
    ShowStatus "Failed to set French language: " & Err.Description
End Sub