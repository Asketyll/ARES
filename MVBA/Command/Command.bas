' Module: Command
' Description: Liste all command
' License: This project is licensed under the AGPL-3.0.
' Dependencies: BootLoader, LangManager, ARESConfigClass, FileDialogs, Zoning, ExportLengthInRegion, CustomPropertyHandler, PropertyRendering, CallStackClass
Option Explicit

Private moZoningGUI          As Zoning_GUI_Options
Private moOutlineGUI         As Outline_GUI_Options
Private moZoneExportGUI      As ExportLengthInReg_GUI_Options
Private moPropertyTaggingGUI As PropertyTagging_GUI_Options
Private moPropertyCalculationGUI     As PropertyCalculation_GUI_Options
Private moPropertyRenderingGUI       As PropertyRendering_GUI_Options

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
' Excel visibility is driven by ARES_Zone_Export_Excel_Visible (default: False;
' user-editable via the "Open once exported" checkbox in EditZoneExportOptions).
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
    Exit Sub

ErrorHandler:
    ReportFailure "SplitRegion", Err.Description, Err.Number, Err.Source
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
    Exit Sub

ErrorHandler:
    ReportFailure "MergeRegion", Err.Description, Err.Number, Err.Source
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
End Sub

' Called from UserForm_QueryClose when form closes
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

' Open the Property Tagging (custom-property) options GUI
Sub EditPropertyTaggingOptions()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If moPropertyTaggingGUI Is Nothing Then
        Set moPropertyTaggingGUI = New PropertyTagging_GUI_Options
    End If

    moPropertyTaggingGUI.Show vbModeless
    ReportIfLogged "EditPropertyTaggingOptions"
    Exit Sub

ErrorHandler:
    ReportFailure "EditPropertyTaggingOptions", Err.Description, Err.Number, Err.Source
End Sub

Public Sub OnPropertyTaggingGUIClosed()
    Set moPropertyTaggingGUI = Nothing
End Sub

' Open the Property Calculation (calc rules -> custom-property values) options GUI
Sub EditPropertyCalculationOptions()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If moPropertyCalculationGUI Is Nothing Then
        Set moPropertyCalculationGUI = New PropertyCalculation_GUI_Options
    End If

    moPropertyCalculationGUI.Show vbModeless
    ReportIfLogged "EditPropertyCalculationOptions"
    Exit Sub

ErrorHandler:
    ReportFailure "EditPropertyCalculationOptions", Err.Description, Err.Number, Err.Source
End Sub

Public Sub OnPropertyCalculationGUIClosed()
    Set moPropertyCalculationGUI = Nothing
End Sub

' Key-in: options panel for Property Rendering - the render master switch plus the three display settings
' that outlive Auto Lengths (colour sync and the two ATLAS label-cell options).
Sub EditPropertyRenderingOptions()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If moPropertyRenderingGUI Is Nothing Then
        Set moPropertyRenderingGUI = New PropertyRendering_GUI_Options
    End If

    moPropertyRenderingGUI.Show vbModeless
    ReportIfLogged "EditPropertyRenderingOptions"
    Exit Sub

ErrorHandler:
    ReportFailure "EditPropertyRenderingOptions", Err.Description, Err.Number, Err.Source
End Sub

Public Sub OnPropertyRenderingGUIClosed()
    Set moPropertyRenderingGUI = Nothing
End Sub

' Key-in: open the DGNLib holding the ARES custom-property ItemTypes, then its Item Types dialog, so the
' definitions (ItemTypes, value lists) can be edited straight away. MicroStation closes the working file
' to do so; re-opening it afterwards refreshes the Item Type state on its own (DGNOpenClose ->
' CustomPropertyHandler.RefreshItemTypes), which closes the edit loop without a restart.
Sub OpenPropertyLibrary()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    ' Library not found (not in MS_DGNLIBLIST, not deployed): an expected user-facing situation, so it is
    ' reported on the status bar only - the technical detail, if any, already went to the log downstream.
    If Not CustomPropertyHandler.OpenCustomPropertyLibrary() Then
        ShowStatusT "PropertyLibraryNotFound"
        Exit Sub
    End If

    ReportIfLogged "OpenPropertyLibrary"
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
    ErrorHandler.ClearErrorFlag
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If Not PropertyRendering.IsEnabled Then
        ShowStatusT "RenderDisabled"
        Exit Sub
    End If

    If Not ActiveModelReference.AnyElementsSelected Then
        ShowStatusT "RenderNoSelection"
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
    ReportIfLogged "BindPropertyRender"
    Exit Sub

ErrorHandler:
    ReportFailure "BindPropertyRender", Err.Description, Err.Number, Err.Source
End Sub

' Key-in: write the deepest/most recent ARES call-stack chain to the log, on demand, without any error
' involved. VBA/MicroStation is single-threaded and synchronous: no ARES procedure is ever still "on the
' stack" by the time a key-in runs (control has already returned to the user), so this cannot read a live
' snapshot - it dumps CallStack.LastSnapshot instead, the chain captured through the most recent ARES
' event/idle pass (see CallStackClass.Push). Reproduces, from a key-in, the same visibility the temporary
' Debug.Print instrumentation used to give during manual debugging sessions.
Sub LogCallStack()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    Dim sSnapshot As String
    sSnapshot = CallStack.LastSnapshot

    If Len(sSnapshot) = 0 Then
        ShowStatusT "CallStackEmpty"
        Exit Sub
    End If

    ErrorHandler.HandleError sSnapshot, 0, "", "Command.LogCallStack"
    ShowStatusT "CallStackLogged"
    ReportIfLogged "LogCallStack"
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
    If Not moPropertyTaggingGUI Is Nothing Then FormPlacement.SaveFormPosition moPropertyTaggingGUI, moPropertyTaggingGUI.Name
    If Not moPropertyCalculationGUI Is Nothing Then FormPlacement.SaveFormPosition moPropertyCalculationGUI, moPropertyCalculationGUI.Name
    If Not moPropertyRenderingGUI Is Nothing Then FormPlacement.SaveFormPosition moPropertyRenderingGUI, moPropertyRenderingGUI.Name
End Sub

' Key-in: forget all saved form positions and re-center any option form currently open.
Sub ResetFormPositions()
    On Error GoTo ErrorHandler
    ErrorHandler.ClearErrorFlag
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If
    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    FormPlacement.ClearFormPositions
    If Not moZoningGUI Is Nothing Then FormPlacement.CenterForm moZoningGUI
    If Not moOutlineGUI Is Nothing Then FormPlacement.CenterForm moOutlineGUI
    If Not moZoneExportGUI Is Nothing Then FormPlacement.CenterForm moZoneExportGUI
    If Not moPropertyTaggingGUI Is Nothing Then FormPlacement.CenterForm moPropertyTaggingGUI
    If Not moPropertyCalculationGUI Is Nothing Then FormPlacement.CenterForm moPropertyCalculationGUI
    If Not moPropertyRenderingGUI Is Nothing Then FormPlacement.CenterForm moPropertyRenderingGUI

    ShowStatusT "FormPositionsReset"
    ReportIfLogged "ResetFormPositions"
    Exit Sub

ErrorHandler:
    ReportFailure "ResetFormPositions", Err.Description, Err.Number, Err.Source
End Sub