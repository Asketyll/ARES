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
    FileDialogs.ExportConfigurationUI
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
    FileDialogs.ImportConfigurationUI
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
    FileDialogs.ShowConfigurationSummaryUI
    Exit Sub
    
ErrorHandler:
    ShowStatus "Configuration summary failed: " & Err.Description
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

' === TESTING COMMANDS ===

' Run all unit tests
Sub RunARESTests()
    On Error GoTo ErrorHandler
    UnitTesting.RunAllTests
    Exit Sub
    
ErrorHandler:
    ShowStatus "Unit tests failed: " & Err.Description
End Sub

' Run performance tests
Sub RunARESPerformanceTests()
    On Error GoTo ErrorHandler
    UnitTesting.RunPerformanceTests
    Exit Sub
    
ErrorHandler:
    ShowStatus "Performance tests failed: " & Err.Description
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

' Sub to open ARES wiki in default browser
Sub OpenARESWiki()
    On Error GoTo ErrorHandler
    
    Dim WikiURL As String
    Dim Result As Long
    
    WikiURL = "https://github.com/Asketyll/ARES/wiki"
    
    ' Use Shell to open URL in default browser
    Result = Shell("rundll32.exe url.dll,FileProtocolHandler " & WikiURL, vbNormalFocus)
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Command.OpenARESWiki"
    If LangManager.IsInit Then
        ShowStatus GetTranslation("WikiOpenFailed") & ": " & Err.Description
    Else
        ShowStatus "Failed to open ARES wiki: " & Err.Description
    End If
End Sub