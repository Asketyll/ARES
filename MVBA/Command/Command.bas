' Module: Command
' Description: Liste all command
' License: This project is licensed under the AGPL-3.0.
' Dependencies: AutoLengths, BootLoader, LangManager, ARESConfigClass, ConfigurationUI, Zoning, ExportLengthInRegion
Option Explicit

Private moAutoLengthsGUI  As AutoLengths_GUI_Options
Private moZoningGUI       As Zoning_GUI_Options
Private moZoneExportGUI   As ExportLengthInReg_GUI_Options

' === AUTO LENGTHS COMMANDS ===

' Sub to call CommandState for manual update length in string
Sub ForceUpdateLength()
    On Error GoTo ErrorHandler
    If Not LicenseManager.IsLicenseValid() Then
        ShowStatus "ARES: License not valid — ForceUpdateLength disabled"
        Exit Sub
    End If
    CommandState.StartLocate New AutoLengths
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Command.ForceUpdateLength"
End Sub

' === UPDATE COMMANDS ===

' Manually check for an available update — bypasses mute and ignore-version preferences
Sub CheckForUpdate()
    On Error GoTo ErrorHandler
    UpdateChecker.CheckForUpdateManual
    Exit Sub

ErrorHandler:
    ShowStatus "Update check failed: " & Err.Description
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
    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    MsgBox ARESConfig.GetConfigSummary(), vbOKOnly + vbInformation, GetTranslation("ConfigSummaryTitle")
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
    If Not LicenseManager.IsLicenseValid() Then
        ShowStatus "ARES: License not valid — EditAutoLengthsOptions disabled"
        Exit Sub
    End If

    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If
    
    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    
    ' Create form only if it doesn't exist
    If moAutoLengthsGUI Is Nothing Then
        Set moAutoLengthsGUI = New AutoLengths_GUI_Options
    End If
    
    ' Show will bring to front if already visible
    moAutoLengthsGUI.Show vbModeless
    
    Exit Sub
    
ErrorHandler:
    ShowStatus "Failed to open AutoLengths options: " & Err.Description
End Sub

' === ZONING COMMANDS ===

' Run zoning using configuration defaults (levels, distance, output properties from ARESConfig)
Sub RunZoning()
    On Error GoTo ErrorHandler
    If Not LicenseManager.IsLicenseValid() Then
        ShowStatus "ARES: License not valid — RunZoning disabled"
        Exit Sub
    End If

    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    Zoning.Zoning
    Exit Sub

ErrorHandler:
    ShowStatus "Zoning failed: " & Err.Description
End Sub

' Run a second, tighter zoning pass: buffer distance from ARES_Zoning2_Distance
' (default 0.2 m), flat (square) caps, per-element sub-zones fused but zones from
' different elements NOT merged.
Sub RunZoning2()
    On Error GoTo ErrorHandler
    If Not LicenseManager.IsLicenseValid() Then
        ShowStatus "ARES: License not valid — RunZoning2 disabled"
        Exit Sub
    End If

    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    ' Resolve the tight buffer distance from its config var. Abort cleanly on an
    ' invalid (<= 0 / empty / non-numeric) value instead of letting the engine
    ' silently fall back to ARES_ZONING_DISTANCE (2.0 m) via its Dist<=0 contract.
    Dim dDist As Double
    dDist = Val(ARESConfig.ARES_ZONING2_DISTANCE.Value)
    If dDist <= 0 Then
        ShowStatus "ARES: ARES_Zoning2_Distance invalid or empty — RunZoning2 aborted"
        Exit Sub
    End If

    Zoning.Zoning Dist:=dDist, MergeZones:=False, RoundCaps:=False
    Exit Sub

ErrorHandler:
    ShowStatus "RunZoning2 failed: " & Err.Description
End Sub

' Export element lengths per zone to Excel.
' Filepath defaults to the active design file's folder (timestamped .xlsx).
' Excel visibility is driven by ARES_Zone_Export_Excel_Visible (default: True).
Sub ExportLength()
    On Error GoTo ErrorHandler
    If Not LicenseManager.IsLicenseValid() Then
        ShowStatus "ARES: License not valid — RunZoneExport disabled"
        Exit Sub
    End If

    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    Dim bVisible As Boolean
    bVisible = (UCase(Trim(ARESConfig.ARES_ZONE_EXPORT_EXCEL_VISIBLE.Value)) = "TRUE")

    ExportLengthInRegion.ExportLengthInRegion ExcelVisible:=bVisible
    Exit Sub

ErrorHandler:
    ShowStatus "ExportLengthInRegion failed: " & Err.Description
End Sub

' Open the Zoning options GUI
Sub EditZoningOptions()
    On Error GoTo ErrorHandler
    If Not LicenseManager.IsLicenseValid() Then
        ShowStatus "ARES: License not valid — EditZoningOptions disabled"
        Exit Sub
    End If

    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If moZoningGUI Is Nothing Then
        Set moZoningGUI = New Zoning_GUI_Options
    End If

    moZoningGUI.Show vbModeless
    Exit Sub

ErrorHandler:
    ShowStatus "Failed to open Zoning options: " & Err.Description
End Sub

' === REGION SPLIT COMMANDS ===

' Split a closed region (Shape / ComplexShape) into two regions with a single datapoint
' on its boundary. The cut runs perpendicular to the local boundary segment at the clicked
' point, across the interior to the opposite boundary. Both halves inherit the original's
' level + symbology; the original is deleted (default) or kept (ARES_RegionSplit_Keep_Original).
Sub SplitRegion()
    On Error GoTo ErrorHandler
    If Not LicenseManager.IsLicenseValid() Then
        ShowStatus "ARES: License not valid — SplitRegion disabled"
        Exit Sub
    End If

    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    CommandState.StartPrimitive New RegionSplitLocate
    Exit Sub

ErrorHandler:
    ShowStatus "SplitRegion failed: " & Err.Description
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

' Called from UserForm_QueryClose when form closes
Public Sub OnAutoLengthsGUIClosed()
    Set moAutoLengthsGUI = Nothing
End Sub

Public Sub OnZoningGUIClosed()
    Set moZoningGUI = Nothing
End Sub

' Open the ZoneExport options GUI
Sub EditZoneExportOptions()
    On Error GoTo ErrorHandler
    If Not LicenseManager.IsLicenseValid() Then
        ShowStatus "ARES: License not valid — EditZoneExportOptions disabled"
        Exit Sub
    End If

    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If moZoneExportGUI Is Nothing Then
        Set moZoneExportGUI = New ExportLengthInReg_GUI_Options
    End If

    moZoneExportGUI.Show vbModeless
    Exit Sub

ErrorHandler:
    ShowStatus "Failed to open ZoneExport options: " & Err.Description
End Sub

Public Sub OnZoneExportGUIClosed()
    Set moZoneExportGUI = Nothing
End Sub

