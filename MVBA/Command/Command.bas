' Module: Command
' Description: Liste all command
' License: This project is licensed under the AGPL-3.0.
' Dependencies: AutoLengths, BootLoader, LangManager, ARESConfigClass, ConfigurationUI, Zoning, ExportLengthInRegion
Option Explicit

Private moAutoLengthsGUI  As AutoLengths_GUI_Options
Private moZoningGUI       As Zoning_GUI_Options
Private moOutlineGUI      As Outline_GUI_Options
Private moZoneExportGUI   As ExportLengthInReg_GUI_Options

' Report a trapped fault from a key-in entry point (messaging rules): log the technical detail
' to the .log (English, via HandleError), then show the user a translated, GENERIC failure line.
' Raw Err.Description never reaches the status bar. Capture Err.* at the handler and pass them in.
Private Sub ReportFailure(ByVal sOp As String, ByVal sDesc As String, ByVal lNum As Long, ByVal sSrc As String)
    On Error Resume Next
    ErrorHandler.HandleError sDesc, lNum, sSrc, "Command." & sOp
    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    ShowStatus GetTranslation("CommandFailed", sOp)
End Sub

' Success-path counterpart to ReportFailure: if a real fault was logged (by this command or a module
' it called) since ClearErrorFlag, tell the user once — with the command's own name. Covers the
' log-and-swallow case where the fault was caught downstream and never reached the ErrorHandler.
Private Sub ReportIfLogged(ByVal sOp As String)
    On Error Resume Next
    If ErrorHandler.HadError Then
        If Not LangManager.IsInit Then LangManager.InitializeTranslations
        ShowStatus GetTranslation("CommandFailed", sOp)
    End If
End Sub

' === AUTO LENGTHS COMMANDS ===

' Sub to call CommandState for manual update length in string
Sub ForceUpdateLength()
    On Error GoTo ErrorHandler
    CommandState.StartLocate New AutoLengths
    Exit Sub

ErrorHandler:
    ReportFailure "ForceUpdateLength", Err.Description, Err.Number, Err.Source
End Sub

' === UPDATE COMMANDS ===

' Manually check for an available update — bypasses mute and ignore-version preferences
Sub CheckForUpdate()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    UpdateChecker.CheckForUpdateManual
    ReportIfLogged "CheckForUpdate"
    Exit Sub

ErrorHandler:
    ReportFailure "CheckForUpdate", Err.Description, Err.Number, Err.Source
End Sub

' === CONFIGURATION MANAGEMENT COMMANDS ===

' Export current configuration using event-driven UI
Sub ExportARESConfig()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    FileDialogs.ExportConfigurationUI
    ReportIfLogged "ExportARESConfig"
    Exit Sub
    
ErrorHandler:
    ReportFailure "ExportARESConfig", Err.Description, Err.Number, Err.Source
End Sub

' Import configuration using event-driven UI
Sub ImportARESConfig()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    FileDialogs.ImportConfigurationUI
    ReportIfLogged "ImportARESConfig"
    Exit Sub
    
ErrorHandler:
    ReportFailure "ImportARESConfig", Err.Description, Err.Number, Err.Source
End Sub

' Show current configuration summary
Sub ShowARESConfigSummary()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    MsgBox ARESConfig.GetConfigSummary(), vbOKOnly + vbInformation, GetTranslation("ConfigSummaryTitle")
    ReportIfLogged "ShowARESConfigSummary"
    Exit Sub

ErrorHandler:
    ReportFailure "ShowARESConfigSummary", Err.Description, Err.Number, Err.Source
End Sub

' === VARIABLE MANAGEMENT COMMANDS ===

' Sub to reset all ARES var in MS
Sub ResetARESVariables()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    
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
    ReportIfLogged "ResetARESVariables"
    
    Exit Sub
    
ErrorHandler:
    ReportFailure "ResetARESVariables", Err.Description, Err.Number, Err.Source
End Sub

' Sub to remove all ARES var in MS
Sub RemoveARESVariables()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    
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
    ReportIfLogged "RemoveARESVariables"
    
    Exit Sub
    
ErrorHandler:
    ReportFailure "RemoveARESVariables", Err.Description, Err.Number, Err.Source
End Sub

' === GUI COMMANDS ===

' Sub to call GUI Options of AutoLengths
Sub EditAutoLengthsOptions()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
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
    ReportIfLogged "EditAutoLengthsOptions"
    
    Exit Sub
    
ErrorHandler:
    ReportFailure "EditAutoLengthsOptions", Err.Description, Err.Number, Err.Source
End Sub

' === ZONING COMMANDS ===

' Run zoning using configuration defaults (levels, distance, output properties from ARESConfig)
Sub RunZoning()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    Zoning.Zoning
    ReportIfLogged "RunZoning"
    Exit Sub

ErrorHandler:
    ReportFailure "RunZoning", Err.Description, Err.Number, Err.Source
End Sub

' Run the Outline pass: a tighter per-element zoning variant driven entirely by its
' own option set (ARES_Outline_* — source levels, distance, output symbology). Flat
' (square) caps, per-element sub-zones fused but zones from different elements NOT
' merged. Edit its options via EditOutlineOptions.
Sub RunOutline()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    ' Resolve Outline's own buffer distance. Abort cleanly on an invalid
    ' (<= 0 / empty / non-numeric) value instead of letting the engine silently
    ' fall back to ARES_ZONING_DISTANCE (2.0 m) via its Dist<=0 contract.
    Dim dDist As Double
    dDist = Val(ARESConfig.ARES_OUTLINE_DISTANCE.Value)
    If dDist <= 0 Then
        ShowStatus GetTranslation("OutlineDistanceInvalid")
        Exit Sub
    End If

    ' Resolve Outline's own source levels. Pass an explicit array so the engine does
    ' not fall back to ARES_ZONING_LEVEL (an empty string would trigger that contract).
    Dim sLvls As String
    sLvls = ARESConfig.ARES_OUTLINE_LEVEL.Value
    If Len(Trim(sLvls)) = 0 Then
        ShowStatus GetTranslation("OutlineLevelEmpty")
        Exit Sub
    End If

    ' Drive the engine from Outline's own option set (output symbology included).
    Zoning.Zoning Lvls:=Split(sLvls, ARES_VAR_DELIMITER), _
                  OutputLevel:=ARESConfig.ARES_OUTLINE_OUTPUT_LEVEL.Value, _
                  Color:=CLng(ARESConfig.ARES_OUTLINE_OUTPUT_COLOR.Value), _
                  Style:=ARESConfig.ARES_OUTLINE_OUTPUT_STYLE.Value, _
                  Weight:=CLng(ARESConfig.ARES_OUTLINE_OUTPUT_WEIGHT.Value), _
                  Dist:=dDist, MergeZones:=False, RoundCaps:=False
    ReportIfLogged "RunOutline"
    Exit Sub

ErrorHandler:
    ReportFailure "RunOutline", Err.Description, Err.Number, Err.Source
End Sub

' Export element lengths per zone to Excel.
' Filepath defaults to the active design file's folder (timestamped .xlsx).
' Excel visibility is driven by ARES_Zone_Export_Excel_Visible (default: True).
Sub ExportLength()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    Dim bVisible As Boolean
    bVisible = (UCase(Trim(ARESConfig.ARES_ZONE_EXPORT_EXCEL_VISIBLE.Value)) = "TRUE")

    ExportLengthInRegion.ExportLengthInRegion ExcelVisible:=bVisible
    ReportIfLogged "ExportLength"
    Exit Sub

ErrorHandler:
    ReportFailure "ExportLength", Err.Description, Err.Number, Err.Source
End Sub

' Open the Zoning options GUI
Sub EditZoningOptions()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If moZoningGUI Is Nothing Then
        Set moZoningGUI = New Zoning_GUI_Options
    End If

    moZoningGUI.Show vbModeless
    ReportIfLogged "EditZoningOptions"
    Exit Sub

ErrorHandler:
    ReportFailure "EditZoningOptions", Err.Description, Err.Number, Err.Source
End Sub

' Open the Outline options GUI
Sub EditOutlineOptions()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If moOutlineGUI Is Nothing Then
        Set moOutlineGUI = New Outline_GUI_Options
    End If

    moOutlineGUI.Show vbModeless
    ReportIfLogged "EditOutlineOptions"
    Exit Sub

ErrorHandler:
    ReportFailure "EditOutlineOptions", Err.Description, Err.Number, Err.Source
End Sub

' === REGION SPLIT COMMANDS ===

' Split a closed region (Shape / ComplexShape) into two regions with a single datapoint
' on its boundary. The cut runs perpendicular to the local boundary segment at the clicked
' point, across the interior to the opposite boundary. Both halves inherit the original's
' level + symbology; the original is deleted (default) or kept (ARES_RegionSplit_Keep_Original).
Sub SplitRegion()
    On Error GoTo ErrorHandler
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    CommandState.StartPrimitive New RegionSplitLocate
    Exit Sub

ErrorHandler:
    ReportFailure "SplitRegion", Err.Description, Err.Number, Err.Source
End Sub

' === TESTING COMMANDS ===

' Run all unit tests
Sub RunARESTests()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    UnitTesting.RunAllTests
    ReportIfLogged "RunARESTests"
    Exit Sub
    
ErrorHandler:
    ReportFailure "RunARESTests", Err.Description, Err.Number, Err.Source
End Sub

' Run performance tests
Sub RunARESPerformanceTests()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    UnitTesting.RunPerformanceTests
    ReportIfLogged "RunARESPerformanceTests"
    Exit Sub
    
ErrorHandler:
    ReportFailure "RunARESPerformanceTests", Err.Description, Err.Number, Err.Source
End Sub

' === LANGUAGE COMMANDS ===

' Sub to set language to English
Sub English()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    
    If Config.SetVar("ARES_Language", "English") Then
        LangManager.InitializeTranslations          ' reload so the confirmation shows in the resolved language
        LangManager.ShowStatusT "LanguageChanged"
    Else
        LangManager.ShowStatusT "LanguageChangeFailed"
    End If
    ReportIfLogged "English"

    Exit Sub

ErrorHandler:
    ReportFailure "English", Err.Description, Err.Number, Err.Source
End Sub

' Sub to set language to French
Sub Français()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    
    If Config.SetVar("ARES_Language", "Français") Then
        LangManager.InitializeTranslations          ' reload so the confirmation shows in the resolved language
        LangManager.ShowStatusT "LanguageChanged"
    Else
        LangManager.ShowStatusT "LanguageChangeFailed"
    End If
    ReportIfLogged "Français"

    Exit Sub

ErrorHandler:
    ReportFailure "Français", Err.Description, Err.Number, Err.Source
End Sub

' Sub to open ARES wiki in default browser
Sub OpenARESWiki()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    
    Dim WikiURL As String
    Dim Result As Long

    ' Open the wiki landing page matching the user's ARES language
    If UCase(Left(LangManager.UserLanguage, 2)) = "FR" Then
        WikiURL = "https://github.com/Asketyll/ARES/wiki/Accueil"
    Else
        WikiURL = "https://github.com/Asketyll/ARES/wiki"
    End If

    ' Use Shell to open URL in default browser
    Result = Shell("rundll32.exe url.dll,FileProtocolHandler " & WikiURL, vbNormalFocus)
    ReportIfLogged "OpenARESWiki"
    
    Exit Sub
    
ErrorHandler:
    ReportFailure "OpenARESWiki", Err.Description, Err.Number, Err.Source
End Sub

' Called from UserForm_QueryClose when form closes
Public Sub OnAutoLengthsGUIClosed()
    Set moAutoLengthsGUI = Nothing
End Sub

Public Sub OnZoningGUIClosed()
    Set moZoningGUI = Nothing
End Sub

Public Sub OnOutlineGUIClosed()
    Set moOutlineGUI = Nothing
End Sub

' Open the ZoneExport options GUI
Sub EditZoneExportOptions()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If moZoneExportGUI Is Nothing Then
        Set moZoneExportGUI = New ExportLengthInReg_GUI_Options
    End If

    moZoneExportGUI.Show vbModeless
    ReportIfLogged "EditZoneExportOptions"
    Exit Sub

ErrorHandler:
    ReportFailure "EditZoneExportOptions", Err.Description, Err.Number, Err.Source
End Sub

Public Sub OnZoneExportGUIClosed()
    Set moZoneExportGUI = Nothing
End Sub

