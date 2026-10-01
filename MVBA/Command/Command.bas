' Module: Command
' Description: Liste all command
' License: This project is licensed under the AGPL-3.0.
' Dependencies: BootLoader, LangManager, ARESConfigClass, FileDialogs, Zoning, ExportLengthInRegion, CustomPropertyHandler, PropertyRendering, CallStackClass, CableReport, SheetLevels, ConfigThemes
Option Explicit

Private moZoningGUI          As Zoning_GUI_Options
Private moOutlineGUI         As Outline_GUI_Options
Private moZoneExportGUI      As ExportLengthInReg_GUI_Options
Private moCableReportGUI     As CableReport_GUI_Options
Private moPropertyTaggingGUI As PropertyTagging_GUI_Options
Private moPropertyCalculationGUI     As PropertyCalculation_GUI_Options
Private moPropertyRenderingGUI       As PropertyRendering_GUI_Options
Private moSheetLevelsGUI     As SheetLevels_GUI_Options
Private moConfigThemesGUI    As ConfigThemes_GUI
Private mbKeyinForm          As Boolean     ' the running key-in opens a form (set by BeginKeyin)
Private msLastKeyin          As String      ' Sub name of the last key-in that wrote the command area

' Report a trapped fault from a key-in entry point (messaging rules): log the technical detail
' to the .log (English, via HandleError), then show the user a translated, GENERIC failure line.
' Raw Err.Description never reaches the status bar. Capture Err.* at the handler and pass them in.
' bAnnounce: leave the name + "Failed" in the command/prompt areas; False for the interactive key-ins,
' whose command name is CommandState.CommandName.
Private Sub ReportFailure(ByVal sOp As String, ByVal sDesc As String, ByVal lNum As Long, ByVal sSrc As String, _
                          Optional ByVal bAnnounce As Boolean = True)
    On Error Resume Next
    ErrorHandler.HandleError sDesc, lNum, sSrc, "Command." & sOp
    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    LangManager.ShowStatusText GetTranslation("CommandFailed", sOp)
    If bAnnounce Then
        ShowCommand KeyinName(sOp)
        ShowPrompt GetTranslation("KeyinFailed")
    End If
End Sub

' Success-path counterpart to ReportFailure: if a real fault was logged (by this command or a module
' it called) since ClearErrorFlag, tell the user once — with the command's own name. Covers the
' log-and-swallow case where the fault was caught downstream and never reached the ErrorHandler.
Private Sub ReportIfLogged(ByVal sOp As String)
    On Error Resume Next
    If ErrorHandler.HadError Then
        If Not LangManager.IsInit Then LangManager.InitializeTranslations
        LangManager.ShowStatusText GetTranslation("CommandFailed", sOp)
    End If
End Sub

' Entry ritual of every synchronous key-in: reset the fault flag (unless bClearFlag is False), then
' announce the key-in's translated name in the command area and what it is doing in the prompt area -
' both transient, never in the Message Center. bForm: the key-in opens an options window.
' English/Francais call it after the language reload with bClearFlag:=False (they clear the flag at
' their top), so the name shows in the new language.
Private Sub BeginKeyin(ByVal sOp As String, Optional ByVal bForm As Boolean = False, _
                       Optional ByVal bClearFlag As Boolean = True)
    On Error Resume Next
    If bClearFlag Then ErrorHandler.ClearErrorFlag
    mbKeyinForm = bForm
    msLastKeyin = sOp
    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    ShowCommand KeyinName(sOp)
    If bForm Then
        ShowPrompt GetTranslation("KeyinFormOpen")
    Else
        ShowPrompt GetTranslation("KeyinWorking")
    End If
End Sub

' Exit ritual, every non-fault return path of a synchronous key-in: report a fault swallowed downstream,
' then leave the name in the command area and the end state in the prompt area. bAchieved: False when
' the key-in refused, was cancelled, or otherwise did not do its job.
Private Sub EndKeyin(ByVal sOp As String, Optional ByVal bAchieved As Boolean = True)
    On Error Resume Next
    ReportIfLogged sOp
    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    ShowCommand KeyinName(sOp)
    ShowPrompt GetTranslation(KeyinEndPromptKey(ErrorHandler.HadError, bAchieved, mbKeyinForm))
End Sub

' Translation key of the prompt a key-in ends on. A fault or an unachieved run wins over everything,
' then a form opener keeps its form prompt, else Done.
Public Function KeyinEndPromptKey(ByVal bHadError As Boolean, ByVal bAchieved As Boolean, _
                                  ByVal bForm As Boolean) As String
    If bHadError Or Not bAchieved Then
        KeyinEndPromptKey = "KeyinFailed"
    ElseIf bForm Then
        KeyinEndPromptKey = "KeyinFormOpen"
    Else
        KeyinEndPromptKey = "KeyinDone"
    End If
End Function

' Translated name of a key-in; the Sub name itself when it has no key. Français is keyed as French:
' translation keys stay ASCII.
Private Function KeyinName(ByVal sOp As String) As String
    Dim sKey As String
    sKey = "Keyin_" & IIf(sOp = "Français", "French", sOp)
    If LangManager.HasTranslation("EN", sKey) Then
        KeyinName = GetTranslation(sKey)
    Else
        KeyinName = sOp
    End If
End Function

' An options window closed: clear the command and prompt areas, but only when its opener is still the
' last key-in shown there - a later key-in, another form or a locate tool keeps what it wrote.
Private Sub ClearKeyinIfLast(ByVal sOpener As String)
    On Error Resume Next
    If msLastKeyin <> sOpener Then Exit Sub
    ShowCommand ""
    ShowPrompt ""
    msLastKeyin = ""
End Sub

' === UPDATE COMMANDS ===

' Manually check for an available update — bypasses mute and ignore-version preferences
Sub CheckForUpdate()
    On Error GoTo ErrorHandler
    BeginKeyin "CheckForUpdate"
    EndKeyin "CheckForUpdate", UpdateChecker.CheckForUpdateManual()
    Exit Sub

ErrorHandler:
    ReportFailure "CheckForUpdate", Err.Description, Err.Number, Err.Source
End Sub

' === CONFIGURATION MANAGEMENT COMMANDS ===

' Export current configuration using event-driven UI
Sub ExportARESConfig()
    On Error GoTo ErrorHandler
    BeginKeyin "ExportARESConfig"
    EndKeyin "ExportARESConfig", FileDialogs.ExportConfigurationUI()
    Exit Sub
    
ErrorHandler:
    ReportFailure "ExportARESConfig", Err.Description, Err.Number, Err.Source
End Sub

' Import configuration using event-driven UI
Sub ImportARESConfig()
    On Error GoTo ErrorHandler
    BeginKeyin "ImportARESConfig"
    EndKeyin "ImportARESConfig", FileDialogs.ImportConfigurationUI()
    Exit Sub
    
ErrorHandler:
    ReportFailure "ImportARESConfig", Err.Description, Err.Number, Err.Source
End Sub

' Show current configuration summary
Sub ShowARESConfigSummary()
    On Error GoTo ErrorHandler
    BeginKeyin "ShowARESConfigSummary"
    If Not LangManager.IsInit Then LangManager.InitializeTranslations
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    MsgBox ARESConfig.GetConfigSummary(), vbOKOnly + vbInformation, GetTranslation("ConfigSummaryTitle")
    EndKeyin "ShowARESConfigSummary"
    Exit Sub

ErrorHandler:
    ReportFailure "ShowARESConfigSummary", Err.Description, Err.Number, Err.Source
End Sub

' Key-in: open the theme window - the named business configurations of the theme folder. Picking one loads it,
' and the current settings can be saved as a theme from there.
Sub OpenARESThemes()
    On Error GoTo ErrorHandler
    BeginKeyin "OpenARESThemes", True
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If moConfigThemesGUI Is Nothing Then
        Set moConfigThemesGUI = New ConfigThemes_GUI
    End If

    moConfigThemesGUI.Show vbModeless
    EndKeyin "OpenARESThemes"
    Exit Sub

ErrorHandler:
    ReportFailure "OpenARESThemes", Err.Description, Err.Number, Err.Source
End Sub

Public Sub OnConfigThemesGUIClosed()
    Set moConfigThemesGUI = Nothing
    ClearKeyinIfLast "OpenARESThemes"
End Sub

' Re-seed every open ARES form from the configuration just loaded, discarding any edit in progress.
Public Sub RefreshOpenForms()
    On Error GoTo ErrorHandler
    If Not moZoningGUI Is Nothing Then moZoningGUI.RefreshFromConfig
    If Not moOutlineGUI Is Nothing Then moOutlineGUI.RefreshFromConfig
    If Not moZoneExportGUI Is Nothing Then moZoneExportGUI.RefreshFromConfig
    If Not moCableReportGUI Is Nothing Then moCableReportGUI.RefreshFromConfig
    If Not moPropertyTaggingGUI Is Nothing Then moPropertyTaggingGUI.RefreshFromConfig
    If Not moPropertyCalculationGUI Is Nothing Then moPropertyCalculationGUI.RefreshFromConfig
    If Not moPropertyRenderingGUI Is Nothing Then moPropertyRenderingGUI.RefreshFromConfig
    If Not moSheetLevelsGUI Is Nothing Then moSheetLevelsGUI.RefreshFromConfig
    If Not moConfigThemesGUI Is Nothing Then moConfigThemesGUI.RefreshFromConfig
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Command.RefreshOpenForms"
End Sub

' === VARIABLE MANAGEMENT COMMANDS ===

' Sub to reset all ARES var in MS
Sub ResetARESVariables()
    On Error GoTo ErrorHandler
    BeginKeyin "ResetARESVariables"
    
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If
    
    Dim bDone As Boolean
    bDone = ARESConfig.ResetAllConfigVars()
    If bDone Then
        If Not LangManager.IsInit Then LangManager.InitializeTranslations
        LangManager.ShowStatusText GetTranslation("VarResetAllSuccess")
    Else
        LangManager.ShowStatusText GetTranslation("VarResetAllFailed")
    End If
    EndKeyin "ResetARESVariables", bDone
    
    Exit Sub
    
ErrorHandler:
    ReportFailure "ResetARESVariables", Err.Description, Err.Number, Err.Source
End Sub

' Sub to remove all ARES var in MS
Sub RemoveARESVariables()
    On Error GoTo ErrorHandler
    BeginKeyin "RemoveARESVariables"
    
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If
    
    Dim bDone As Boolean
    bDone = ARESConfig.RemoveAllConfigVars()
    If bDone Then
        If Not LangManager.IsInit Then LangManager.InitializeTranslations
        LangManager.ShowStatusText GetTranslation("VarRemoveSuccess")
    Else
        LangManager.ShowStatusText GetTranslation("VarRemoveError")
    End If
    EndKeyin "RemoveARESVariables", bDone
    
    Exit Sub
    
ErrorHandler:
    ReportFailure "RemoveARESVariables", Err.Description, Err.Number, Err.Source
End Sub

' === GUI COMMANDS ===

' === ZONING COMMANDS ===

' Run zoning using configuration defaults (levels, distance, output properties from ARESConfig)
Sub RunZoning()
    On Error GoTo ErrorHandler
    BeginKeyin "RunZoning"
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    EndKeyin "RunZoning", Zoning.Zoning()
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
    BeginKeyin "RunOutline"
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
        LangManager.ShowStatusText GetTranslation("OutlineDistanceInvalid")
        EndKeyin "RunOutline", False
        Exit Sub
    End If

    ' Resolve Outline's own source levels. Pass an explicit array so the engine does
    ' not fall back to ARES_ZONING_LEVEL (an empty string would trigger that contract).
    Dim sLvls As String
    sLvls = ARESConfig.ARES_OUTLINE_LEVEL.Value
    If Len(Trim(sLvls)) = 0 Then
        LangManager.ShowStatusText GetTranslation("OutlineLevelEmpty")
        EndKeyin "RunOutline", False
        Exit Sub
    End If

    ' Drive the engine from Outline's own option set (output symbology included).
    Dim bDone As Boolean
    bDone = Zoning.Zoning(Lvls:=Split(sLvls, ARES_VAR_DELIMITER), _
                          OutputLevel:=ARESConfig.ARES_OUTLINE_OUTPUT_LEVEL.Value, _
                          Color:=CLng(ARESConfig.ARES_OUTLINE_OUTPUT_COLOR.Value), _
                          Style:=ARESConfig.ARES_OUTLINE_OUTPUT_STYLE.Value, _
                          Weight:=CLng(ARESConfig.ARES_OUTLINE_OUTPUT_WEIGHT.Value), _
                          Dist:=dDist, MergeZones:=False, RoundCaps:=False)
    EndKeyin "RunOutline", bDone
    Exit Sub

ErrorHandler:
    ReportFailure "RunOutline", Err.Description, Err.Number, Err.Source
End Sub

' Export element lengths per zone to Excel.
' Filepath defaults to the active design file's folder (timestamped .xlsx).
' Excel visibility is driven by ARES_Zone_Export_Excel_Visible (default: False;
' user-editable via the "Open once exported" checkbox in EditZoneExportOptions).
Sub ExportLength()
    On Error GoTo ErrorHandler
    BeginKeyin "ExportLength"
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    Dim bVisible As Boolean
    bVisible = (UCase(Trim(ARESConfig.ARES_ZONE_EXPORT_EXCEL_VISIBLE.Value)) = "TRUE")

    EndKeyin "ExportLength", ExportLengthInRegion.ExportLengthInRegion(ExcelVisible:=bVisible)
    Exit Sub

ErrorHandler:
    ReportFailure "ExportLength", Err.Description, Err.Number, Err.Source
End Sub

' Export a cable-by-cable trenching report to Excel: end-cell markers (Repere), Nature (read on the
' cable, then its graphic group), Longueur (off the group), and the trenching length broken down by
' soil type (Coupe_Type), pivoted into one column per distinct value. Aerial cables go on their own
' sheet, never measured. Excel visibility is driven by ARES_CableReport_Excel_Visible.
Sub ExportCableReport()
    On Error GoTo ErrorHandler
    BeginKeyin "ExportCableReport"
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    Dim bVisible As Boolean
    bVisible = (UCase(Trim(ARESConfig.ARES_CABLEREPORT_EXCEL_VISIBLE.Value)) = "TRUE")

    EndKeyin "ExportCableReport", CableReport.CableReport(ExcelVisible:=bVisible)
    Exit Sub

ErrorHandler:
    ReportFailure "ExportCableReport", Err.Description, Err.Number, Err.Source
End Sub

' Open the CableReport options GUI
Sub EditCableReportOptions()
    On Error GoTo ErrorHandler
    BeginKeyin "EditCableReportOptions", True
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If moCableReportGUI Is Nothing Then
        Set moCableReportGUI = New CableReport_GUI_Options
    End If

    moCableReportGUI.Show vbModeless
    EndKeyin "EditCableReportOptions"
    Exit Sub

ErrorHandler:
    ReportFailure "EditCableReportOptions", Err.Description, Err.Number, Err.Source
End Sub

Public Sub OnCableReportGUIClosed()
    Set moCableReportGUI = Nothing
    ClearKeyinIfLast "EditCableReportOptions"
End Sub

' Open the Zoning options GUI
Sub EditZoningOptions()
    On Error GoTo ErrorHandler
    BeginKeyin "EditZoningOptions", True
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If moZoningGUI Is Nothing Then
        Set moZoningGUI = New Zoning_GUI_Options
    End If

    moZoningGUI.Show vbModeless
    EndKeyin "EditZoningOptions"
    Exit Sub

ErrorHandler:
    ReportFailure "EditZoningOptions", Err.Description, Err.Number, Err.Source
End Sub

' Open the Outline options GUI
Sub EditOutlineOptions()
    On Error GoTo ErrorHandler
    BeginKeyin "EditOutlineOptions", True
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If moOutlineGUI Is Nothing Then
        Set moOutlineGUI = New Outline_GUI_Options
    End If

    moOutlineGUI.Show vbModeless
    EndKeyin "EditOutlineOptions"
    Exit Sub

ErrorHandler:
    ReportFailure "EditOutlineOptions", Err.Description, Err.Number, Err.Source
End Sub

' === REGION SPLIT/MERGE COMMANDS ===

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
    CommandState.CommandName = GetTranslation("RegionSplitSelectRegionC")   ' after StartPrimitive: tied to the undo buffer
    msLastKeyin = "SplitRegion"
    Exit Sub

ErrorHandler:
    ReportFailure "SplitRegion", Err.Description, Err.Number, Err.Source, bAnnounce:=False
End Sub

' Merge two closed regions (Shape / ComplexShape) into a single region from two successive
' datapoints. The merged region inherits the FIRST clicked region's level + symbology; both
' originals are deleted (default) or kept (ARES_RegionMerge_Keep_Originals).
Sub MergeRegion()
    On Error GoTo ErrorHandler
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    CommandState.StartPrimitive New RegionMergeLocate
    CommandState.CommandName = GetTranslation("MergeRegionSelectFirstC")   ' after StartPrimitive: tied to the undo buffer
    msLastKeyin = "MergeRegion"
    Exit Sub

ErrorHandler:
    ReportFailure "MergeRegion", Err.Description, Err.Number, Err.Source, bAnnounce:=False
End Sub

' === SHEET LEVELS COMMANDS ===

' Turn Global Display on for every level of every Sheet ("Papier") model of the active design file
' whose name matches ARES_Sheet_Levels_Model_Name (default "*Folio*"; | -separated wildcard patterns,
' case-insensitive). Freezing and per-view display are deliberately left untouched - see the
' SheetLevels module header for why the per-view masks are out of reach from here.
Sub ActivateSheetLevels()
    On Error GoTo ErrorHandler
    BeginKeyin "ActivateSheetLevels"
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    EndKeyin "ActivateSheetLevels", SheetLevels.ActivateLevels()
    Exit Sub

ErrorHandler:
    ReportFailure "ActivateSheetLevels", Err.Description, Err.Number, Err.Source
End Sub

' Key-in: edit the Sheet Levels options - the sheet-model name pattern, and whether the sheets'
' references are processed too.
Sub EditSheetLevelsOptions()
    On Error GoTo ErrorHandler
    BeginKeyin "EditSheetLevelsOptions", True
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If moSheetLevelsGUI Is Nothing Then
        Set moSheetLevelsGUI = New SheetLevels_GUI_Options
    End If

    moSheetLevelsGUI.Show vbModeless
    EndKeyin "EditSheetLevelsOptions"
    Exit Sub

ErrorHandler:
    ReportFailure "EditSheetLevelsOptions", Err.Description, Err.Number, Err.Source
End Sub

Public Sub OnSheetLevelsGUIClosed()
    Set moSheetLevelsGUI = Nothing
    ClearKeyinIfLast "EditSheetLevelsOptions"
End Sub

' === TESTING COMMANDS ===

' Run all unit tests
Sub RunARESTests()
    On Error GoTo ErrorHandler
    BeginKeyin "RunARESTests"
    UnitTesting.RunAllTests
    EndKeyin "RunARESTests"
    Exit Sub
    
ErrorHandler:
    ReportFailure "RunARESTests", Err.Description, Err.Number, Err.Source
End Sub

' Run performance tests
Sub RunARESPerformanceTests()
    On Error GoTo ErrorHandler
    BeginKeyin "RunARESPerformanceTests"
    UnitTesting.RunPerformanceTests
    EndKeyin "RunARESPerformanceTests"
    Exit Sub
    
ErrorHandler:
    ReportFailure "RunARESPerformanceTests", Err.Description, Err.Number, Err.Source
End Sub

' === LANGUAGE COMMANDS ===

' Sub to set language to English
Sub English()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    
    Dim bDone As Boolean
    bDone = Config.SetVar("ARES_Language", "English")
    If bDone Then
        LangManager.InitializeTranslations          ' reload so the confirmation shows in the resolved language
        LangManager.ShowStatusT "LanguageChanged"
    Else
        LangManager.ShowStatusT "LanguageChangeFailed"
    End If
    BeginKeyin "English", bClearFlag:=False            ' announced after the reload: the name is in the new language
    EndKeyin "English", bDone

    Exit Sub

ErrorHandler:
    ReportFailure "English", Err.Description, Err.Number, Err.Source
End Sub

' Sub to set language to French
Sub Français()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    
    Dim bDone As Boolean
    bDone = Config.SetVar("ARES_Language", "Français")
    If bDone Then
        LangManager.InitializeTranslations          ' reload so the confirmation shows in the resolved language
        LangManager.ShowStatusT "LanguageChanged"
    Else
        LangManager.ShowStatusT "LanguageChangeFailed"
    End If
    BeginKeyin "Français", bClearFlag:=False            ' announced after the reload: the name is in the new language
    EndKeyin "Français", bDone

    Exit Sub

ErrorHandler:
    ReportFailure "Français", Err.Description, Err.Number, Err.Source
End Sub

' Sub to open ARES wiki in default browser
Sub OpenARESWiki()
    On Error GoTo ErrorHandler
    BeginKeyin "OpenARESWiki"
    
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
    EndKeyin "OpenARESWiki"
    
    Exit Sub

ErrorHandler:
    ReportFailure "OpenARESWiki", Err.Description, Err.Number, Err.Source
End Sub

' Open a SPECIFIC ARES wiki page in the default browser, resolving EN/FR by the user's ARES language - the
' Property Tagging / Property Calculation options forms' help button uses this (their ComboBox tooltip has
' no room for the full grammar reference; the wiki page is the authoritative source). NOT a key-in itself
' (no ClearErrorFlag/ReportIfLogged ritual - it is called from a form button's own error-handled Click).
Public Sub OpenARESWikiPage(ByVal sEnPage As String, ByVal sFrPage As String)
    On Error GoTo ErrorHandler

    Dim WikiURL As String
    If UCase(Left(LangManager.UserLanguage, 2)) = "FR" Then
        WikiURL = "https://github.com/Asketyll/ARES/wiki/" & sFrPage
    Else
        WikiURL = "https://github.com/Asketyll/ARES/wiki/" & sEnPage
    End If

    Shell "rundll32.exe url.dll,FileProtocolHandler " & WikiURL, vbNormalFocus
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Command.OpenARESWikiPage"
    ' A Help button that does nothing looks broken. Unlike the OpenARESWiki key-in, this path has no
    ' ReportIfLogged to surface the fault, so it says so itself.
    LangManager.ShowStatusT "WikiOpenFailed"
End Sub

' Called from UserForm_QueryClose when form closes
Public Sub OnZoningGUIClosed()
    Set moZoningGUI = Nothing
    ClearKeyinIfLast "EditZoningOptions"
End Sub

Public Sub OnOutlineGUIClosed()
    Set moOutlineGUI = Nothing
    ClearKeyinIfLast "EditOutlineOptions"
End Sub

' Open the ZoneExport options GUI
Sub EditZoneExportOptions()
    On Error GoTo ErrorHandler
    BeginKeyin "EditZoneExportOptions", True
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If moZoneExportGUI Is Nothing Then
        Set moZoneExportGUI = New ExportLengthInReg_GUI_Options
    End If

    moZoneExportGUI.Show vbModeless
    EndKeyin "EditZoneExportOptions"
    Exit Sub

ErrorHandler:
    ReportFailure "EditZoneExportOptions", Err.Description, Err.Number, Err.Source
End Sub

Public Sub OnZoneExportGUIClosed()
    Set moZoneExportGUI = Nothing
    ClearKeyinIfLast "EditZoneExportOptions"
End Sub

' Open the Property Tagging (custom-property) options GUI
Sub EditPropertyTaggingOptions()
    On Error GoTo ErrorHandler
    BeginKeyin "EditPropertyTaggingOptions", True
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If moPropertyTaggingGUI Is Nothing Then
        Set moPropertyTaggingGUI = New PropertyTagging_GUI_Options
    End If

    moPropertyTaggingGUI.Show vbModeless
    EndKeyin "EditPropertyTaggingOptions"
    Exit Sub

ErrorHandler:
    ReportFailure "EditPropertyTaggingOptions", Err.Description, Err.Number, Err.Source
End Sub

Public Sub OnPropertyTaggingGUIClosed()
    Set moPropertyTaggingGUI = Nothing
    ClearKeyinIfLast "EditPropertyTaggingOptions"
End Sub

' Open the Property Calculation (calc rules -> custom-property values) options GUI
Sub EditPropertyCalculationOptions()
    On Error GoTo ErrorHandler
    BeginKeyin "EditPropertyCalculationOptions", True
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If moPropertyCalculationGUI Is Nothing Then
        Set moPropertyCalculationGUI = New PropertyCalculation_GUI_Options
    End If

    moPropertyCalculationGUI.Show vbModeless
    EndKeyin "EditPropertyCalculationOptions"
    Exit Sub

ErrorHandler:
    ReportFailure "EditPropertyCalculationOptions", Err.Description, Err.Number, Err.Source
End Sub

Public Sub OnPropertyCalculationGUIClosed()
    Set moPropertyCalculationGUI = Nothing
    ClearKeyinIfLast "EditPropertyCalculationOptions"
End Sub

' Key-in: options panel for Property Rendering - the render master switch plus the three display settings
' that outlive Auto Lengths (colour sync and the two ATLAS label-cell options).
Sub EditPropertyRenderingOptions()
    On Error GoTo ErrorHandler
    BeginKeyin "EditPropertyRenderingOptions", True
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If moPropertyRenderingGUI Is Nothing Then
        Set moPropertyRenderingGUI = New PropertyRendering_GUI_Options
    End If

    moPropertyRenderingGUI.Show vbModeless
    EndKeyin "EditPropertyRenderingOptions"
    Exit Sub

ErrorHandler:
    ReportFailure "EditPropertyRenderingOptions", Err.Description, Err.Number, Err.Source
End Sub

Public Sub OnPropertyRenderingGUIClosed()
    Set moPropertyRenderingGUI = Nothing
    ClearKeyinIfLast "EditPropertyRenderingOptions"
End Sub

' Key-in: open the DGNLib holding the ARES custom-property ItemTypes, then its Item Types dialog, so the
' definitions (ItemTypes, value lists) can be edited straight away. MicroStation closes the working file
' to do so; re-opening it afterwards refreshes the Item Type state on its own (DGNOpenClose ->
' CustomPropertyHandler.RefreshItemTypes), which closes the edit loop without a restart.
Sub OpenPropertyLibrary()
    On Error GoTo ErrorHandler
    BeginKeyin "OpenPropertyLibrary"
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    ' Library not found (not in MS_DGNLIBLIST, not deployed): an expected user-facing situation, so it is
    ' reported on the status bar only - the technical detail, if any, already went to the log downstream.
    If Not CustomPropertyHandler.OpenCustomPropertyLibrary() Then
        ShowStatusT "PropertyLibraryNotFound"
        EndKeyin "OpenPropertyLibrary", False
        Exit Sub
    End If

    EndKeyin "OpenPropertyLibrary"
    Exit Sub

ErrorHandler:
    ReportFailure "OpenPropertyLibrary", Err.Description, Err.Number, Err.Source
End Sub

' Key-in: link the SELECTED text(s) to the custom properties their "Prop[Name]" tokens name, and render
' them once. This is the manual entry point the hybrid auto-bind deliberately leaves open: automatic
' binding only happens when the token's property is ALREADY attached to the element, so an ungrouped text
' matching no tagging rule - or a text authored before its property was attached - is bound from here.
' Operates on the current selection; an empty selection or a disabled feature is status-only.
Sub BindPropertyRender()
    On Error GoTo ErrorHandler
    BeginKeyin "BindPropertyRender"
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If Not PropertyRendering.IsEnabled Then
        ShowStatusT "RenderDisabled"
        EndKeyin "BindPropertyRender", False
        Exit Sub
    End If

    If Not ActiveModelReference.AnyElementsSelected Then
        ShowStatusT "RenderNoSelection"
        EndKeyin "BindPropertyRender", False
        Exit Sub
    End If

    Dim oEnum As ElementEnumerator
    Dim oEl As element

    Set oEnum = ActiveModelReference.GetSelectedElements
    Do While oEnum.MoveNext
        Set oEl = oEnum.Current
        PropertyRendering.BindElement oEl
    Loop

    ' BindElement already reports every refusal on the status bar (and the success of a real bind), so a
    ' zero count needs no extra message here.
    EndKeyin "BindPropertyRender"
    Exit Sub

ErrorHandler:
    ReportFailure "BindPropertyRender", Err.Description, Err.Number, Err.Source
End Sub

' Key-in: force a full pass (Tagging -> Calculation -> Actuator -> Rendering, then Branch 2's sibling
' propagation) on every selected element, exactly as a native ElementChanged event would - the manual entry
' point for whatever automatic propagation did not reach. Being an independent, non-reentrant top-level
' call, it always runs on a fully-settled model. No feature-specific gate here: each pipeline stage already
' gates itself inside ProcessElement, same as it would for a real event.
Sub RecalculateSelection()
    On Error GoTo ErrorHandler
    BeginKeyin "RecalculateSelection"
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If Not ActiveModelReference.AnyElementsSelected Then
        ShowStatusT "RecalculateNoSelection"
        EndKeyin "RecalculateSelection", False
        Exit Sub
    End If

    Dim oEnum As ElementEnumerator
    Dim oEl As element
    Dim nCount As Long

    Set oEnum = ActiveModelReference.GetSelectedElements
    Do While oEnum.MoveNext
        Set oEl = oEnum.Current
        ChangeHandler.ProcessElement oEl
        nCount = nCount + 1
    Loop

    LangManager.ShowStatusText GetTranslation("RecalculateComplete", nCount)
    EndKeyin "RecalculateSelection"
    Exit Sub

ErrorHandler:
    ReportFailure "RecalculateSelection", Err.Description, Err.Number, Err.Source
End Sub

' Key-in: write the deepest/most recent ARES call-stack chain to the log, on demand, without any error
' involved. VBA/MicroStation is single-threaded and synchronous: no ARES procedure is ever still "on the
' stack" by the time a key-in runs (control has already returned to the user), so this cannot read a live
' snapshot - it dumps CallStack.LastSnapshot instead, the chain captured through the most recent ARES
' event/idle pass (see CallStackClass.Push). Reproduces, from a key-in, the same visibility the temporary
' Debug.Print instrumentation used to give during manual debugging sessions.
Sub LogCallStack()
    On Error GoTo ErrorHandler
    BeginKeyin "LogCallStack"

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    Dim sSnapshot As String
    sSnapshot = CallStack.LastSnapshot

    If Len(sSnapshot) = 0 Then
        ShowStatusT "CallStackEmpty"
        EndKeyin "LogCallStack", False
        Exit Sub
    End If

    ErrorHandler.HandleError sSnapshot, 0, "", "Command.LogCallStack"
    ShowStatusT "CallStackLogged"
    EndKeyin "LogCallStack"
    Exit Sub

ErrorHandler:
    ReportFailure "LogCallStack", Err.Description, Err.Number, Err.Source
End Sub

' Persist the position of every option form still open (best-effort; called at project unload).
Public Sub SaveAllOpenFormPositions()
    On Error Resume Next
    If Not moZoningGUI Is Nothing Then FormPlacement.SaveFormPosition moZoningGUI, moZoningGUI.Name
    If Not moOutlineGUI Is Nothing Then FormPlacement.SaveFormPosition moOutlineGUI, moOutlineGUI.Name
    If Not moZoneExportGUI Is Nothing Then FormPlacement.SaveFormPosition moZoneExportGUI, moZoneExportGUI.Name
    If Not moCableReportGUI Is Nothing Then FormPlacement.SaveFormPosition moCableReportGUI, moCableReportGUI.Name
    If Not moPropertyTaggingGUI Is Nothing Then FormPlacement.SaveFormPosition moPropertyTaggingGUI, moPropertyTaggingGUI.Name
    If Not moPropertyCalculationGUI Is Nothing Then FormPlacement.SaveFormPosition moPropertyCalculationGUI, moPropertyCalculationGUI.Name
    If Not moPropertyRenderingGUI Is Nothing Then FormPlacement.SaveFormPosition moPropertyRenderingGUI, moPropertyRenderingGUI.Name
    If Not moSheetLevelsGUI Is Nothing Then FormPlacement.SaveFormPosition moSheetLevelsGUI, moSheetLevelsGUI.Name
    If Not moConfigThemesGUI Is Nothing Then FormPlacement.SaveFormPosition moConfigThemesGUI, moConfigThemesGUI.Name
End Sub

' Key-in: forget all saved form positions and re-center any option form currently open.
Sub ResetFormPositions()
    On Error GoTo ErrorHandler
    BeginKeyin "ResetFormPositions"
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If
    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    FormPlacement.ClearFormPositions
    If Not moZoningGUI Is Nothing Then FormPlacement.CenterForm moZoningGUI
    If Not moOutlineGUI Is Nothing Then FormPlacement.CenterForm moOutlineGUI
    If Not moZoneExportGUI Is Nothing Then FormPlacement.CenterForm moZoneExportGUI
    If Not moCableReportGUI Is Nothing Then FormPlacement.CenterForm moCableReportGUI
    If Not moPropertyTaggingGUI Is Nothing Then FormPlacement.CenterForm moPropertyTaggingGUI
    If Not moPropertyCalculationGUI Is Nothing Then FormPlacement.CenterForm moPropertyCalculationGUI
    If Not moPropertyRenderingGUI Is Nothing Then FormPlacement.CenterForm moPropertyRenderingGUI
    If Not moSheetLevelsGUI Is Nothing Then FormPlacement.CenterForm moSheetLevelsGUI
    If Not moConfigThemesGUI Is Nothing Then FormPlacement.CenterForm moConfigThemesGUI

    ShowStatusT "FormPositionsReset"
    EndKeyin "ResetFormPositions"
    Exit Sub

ErrorHandler:
    ReportFailure "ResetFormPositions", Err.Description, Err.Number, Err.Source
End Sub
