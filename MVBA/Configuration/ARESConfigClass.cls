' Class Module: ARESConfigClass
' Description: Manages all ARES configuration variables
' License: This project is licensed under the AGPL-3.0.
' Dependencies: Config, ARES_MS_VAR_Class, ARESConstants, ErrorHandlerClass
Option Explicit

Private moConfigVariables As Collection
Private mbInitialized As Boolean

'######################################################################################################################
'                             CAN BE MODIFIED IN MS ENVIRONMENT VARIABLES DO NOT MODIFY HERE
'                                   USE CONFIG MODULE TO GET, SET AND REMOVE A VALUE
'######################################################################################################################

' RETIRED, KEPT ON PURPOSE - no code reads these two any more (their last readers were StringsInEl's
' TRIGGER mode, which lost its callers with Auto Lengths). They survive as a DOCUMENTARY TRACE for the
' manual conversion of old-format drawings: they record the pattern a site used to mark an Auto Lengths
' text, and that value is per-site configuration which cannot be reconstructed from source once deleted.
' DO NOT delete them as dead - they are, and that is not the point. DO NOT expose them in any options
' form either: a control editing them would show a setting with no effect.
' They go when Asketyll declares the manual conversion finished, not before.
' Used in LangManager module to force language if CONNECTUSER_LANGUAGE configuration variable is not set
Public ARES_LANGUAGE As ARES_MS_VAR_Class 'No Default Value
' Used in CellRedreaw module to chek if cell is like label (use cell name)
Public ARES_CELL_LIKE_LABEL As ARES_MS_VAR_Class 'Default Value:ETI076|ETI01E|ETI01F|ETI01K|ETI01M|ETI01N|ETI01O      can a array use | (ARES_VAR_DELIMITER)
' Used in CellRedreaw module for automatically enable cell update
Public ARES_UPDATE_ATLASCELLLABEL As ARES_MS_VAR_Class 'Dafault Value: True          True or False (Boolean)
' Used in ElementChangeHandler for bulk operation detection threshold
Public ARES_BULK_THRESHOLD As ARES_MS_VAR_Class 'Default Value: 10    Number of calls to trigger bulk mode
' Used in ElementChangeHandler for bulk operation detection time window
Public ARES_BULK_INTERVAL As ARES_MS_VAR_Class 'Default Value: 1000   Time window in milliseconds
' Used in UpdateChecker module to permanently mute update notifications
Public ARES_UPDATE_MUTE As ARES_MS_VAR_Class 'Default Value: False         True or False (Boolean)
' Used in UpdateChecker module to ignore a specific release version
Public ARES_UPDATE_IGNORE_VERSION As ARES_MS_VAR_Class 'Default Value: (empty)   Version string e.g. "1.2.3"
' Used in Zoning module for source level names to process
Public ARES_ZONING_LEVEL As ARES_MS_VAR_Class 'Default Value: (empty)       can a array use | (ARES_VAR_DELIMITER)
' Used in Zoning module for output level name
Public ARES_ZONING_OUTPUT_LEVEL As ARES_MS_VAR_Class 'Default Value: ARES_Zoning   Level name string
' Used in Zoning module for output element color
Public ARES_ZONING_OUTPUT_COLOR As ARES_MS_VAR_Class 'Default Value: 5             MicroStation color index (Long)
' Used in Zoning module for output element line style
Public ARES_ZONING_OUTPUT_STYLE As ARES_MS_VAR_Class 'Default Value: 0             MicroStation line style index (Long)
' Used in Zoning module for output element line weight
Public ARES_ZONING_OUTPUT_WEIGHT As ARES_MS_VAR_Class 'Default Value: 4            MicroStation line weight index (Long)
' Used in Zoning module for default buffer distance
Public ARES_ZONING_DISTANCE As ARES_MS_VAR_Class 'Default Value: 2.0              Buffer distance in master units (Double)
' === Outline (tight per-element zoning variant) — own option set ===
' Used in Outline module for source level names to process
Public ARES_OUTLINE_LEVEL As ARES_MS_VAR_Class 'Default Value: (empty)          can a array use | (ARES_VAR_DELIMITER)
' Used in Outline module for default buffer distance
Public ARES_OUTLINE_DISTANCE As ARES_MS_VAR_Class 'Default Value: 0.2            Buffer distance in master units (Double)
' Used in Outline module for output level name
Public ARES_OUTLINE_OUTPUT_LEVEL As ARES_MS_VAR_Class 'Default Value: ARES_Outline   Level name string
' Used in Outline module for output element color
Public ARES_OUTLINE_OUTPUT_COLOR As ARES_MS_VAR_Class 'Default Value: 5           MicroStation color index (Long)
' Used in Outline module for output element line style
Public ARES_OUTLINE_OUTPUT_STYLE As ARES_MS_VAR_Class 'Default Value: 0           MicroStation line style index (Long)
' Used in Outline module for output element line weight
Public ARES_OUTLINE_OUTPUT_WEIGHT As ARES_MS_VAR_Class 'Default Value: 4          MicroStation line weight index (Long)
' Used in ExportLengthInRegion module — open the workbook in Excel after export when True; False = headless silent export
Public ARES_ZONE_EXPORT_EXCEL_VISIBLE As ARES_MS_VAR_Class 'Default Value: false          True or False (Boolean)
' Used in ExportLengthInRegion module — name of the worksheet tab in the output Excel workbook
Public ARES_ZONE_EXPORT_SHEET_NAME As ARES_MS_VAR_Class 'Default Value: Export                      String
' Used in ExportLengthInRegion module — True = prompt user with a Save dialog; False = auto-generate timestamped path
Public ARES_ZONE_EXPORT_USE_DIALOG As ARES_MS_VAR_Class 'Default Value: True                        True or False (Boolean)
' Used in ExportLengthInRegion module — decimal places applied to exported lengths; 255 is reserved for error (ARES_RND_ERROR_VALUE)
Public ARES_ZONE_EXPORT_ROUND As ARES_MS_VAR_Class 'Default Value: 2                               Range 0 to 254 (Byte -1) #255 it reserved for error (ARES_RND_ERROR_VALUE)
' Used in ExportLengthInRegion module — grouping key for exported rows: Style = line style name, Level = element level name, Color = color index
Public ARES_ZONE_EXPORT_GROUP_BY As ARES_MS_VAR_Class 'Default Value: Style                        "Style" or "Level" or "Color"
' Used in ExportLengthInRegion module — | -delimited list of level names restricting the measured candidate elements. Empty = all levels (no filter).
Public ARES_ZONE_EXPORT_LEVEL As ARES_MS_VAR_Class 'Default Value: (empty)   |-delimited level list (ARES_VAR_DELIMITER); empty = all levels
' Used in ExportLengthInRegion module — True = additionally split every group by zone (long format: Zone | group | length); False = classic global export.
Public ARES_ZONE_EXPORT_PER_ZONE As ARES_MS_VAR_Class 'Default Value: False                       True or False (Boolean)
' Used in ExportLengthInRegion module — custom-property NAME whose value labels each zone (read on the ZONE element) when per-zone split is on.
' Must be an ItemType of the "ARES" DGNLib. Empty/invalid = zones labeled "Zone <n>" (scan order).
Public ARES_ZONE_EXPORT_ZONE_PROPERTY As ARES_MS_VAR_Class 'Default Value: (empty)   custom-property name; must be an ItemType of the "ARES" DGNLib
' Used in RegionSplit module — geometric epsilon: degenerate-segment / zero-parameter / coincident-point threshold; also drives the cut-knife half-width
Public ARES_REGIONSPLIT_COLLINEAR_TOL As ARES_MS_VAR_Class 'Default Value: 0.000001               Geometric epsilon in master units (Double)
' Used in RegionSplit module — False = original region is deleted after the two halves are written; True = original is kept alongside the two halves
Public ARES_REGIONSPLIT_KEEP_ORIGINAL As ARES_MS_VAR_Class 'Default Value: False                  True or False (Boolean)
' Used in RegionSplit module — max chordal deviation (master units) when densifying an arc boundary side into a polyline; smaller => more chords => closer-to-radial perpendicular and tighter entry foot
Public ARES_REGIONSPLIT_STROKE_TOL As ARES_MS_VAR_Class 'Default Value: 0.01                       Stroke chordal tolerance in master units (Double, > 0)
' Used in RegionMerge module — False = both originals are deleted after the merged region is written; True = both originals are kept alongside the merged result
Public ARES_REGIONMERGE_KEEP_ORIGINALS As ARES_MS_VAR_Class 'Default Value: False                  True or False (Boolean)
' Used in PropertyTagging module — master switch for auto-attaching custom properties on element add/modify.
Public ARES_AUTO_PROPERTIES As ARES_MS_VAR_Class 'Default Value: True                       True or False (Boolean)
' Used in PropertyTagging module — grammar v2 rules "[@]cond[&cond]*=prop|prop ; …" mapping elements to properties to attach
' (cond = [!]Lvl|Cell|Type[name|name]; & = AND; ! negates; */? wildcards; a leading @ attaches to the OTHER group members).
Public ARES_PROPERTY_RULES As ARES_MS_VAR_Class 'Default Value: (empty)                     e.g. Lvl[WALLS]=Commune|Coupe_Type ; @Cell[ETI0*]=Commune
' Used in PropertyCalculation module — calc rules "Prop[name][&conditions]=Source ; …" giving each custom property its computed/fixed value (Source = GroupCellText[pattern]/Value[text]/Coord[n]/Id).
Public ARES_CALC_RULES As ARES_MS_VAR_Class 'Default Value: (empty)                        e.g. Prop[Repere]=GroupCellText[ETI*] ; Prop[XY]=Coord
' Used in PropertyCalculation module — master switch for the value-calculation engine: each custom property gets its value from the first matching ARES_Calc_Rules rule, written ONLY where the property is already attached (attaching stays PropertyTagging's job).
Public ARES_PROPERTY_CALC As ARES_MS_VAR_Class 'Default Value: False                     True or False (Boolean)
' Used in PropertyCalculation module — when True, an emptied value detaches the property (via the tagger, PropertyTagging.DetachRuleProperty) instead of clearing it; a rule that still mandates the property re-attaches it empty.
Public ARES_CALC_DETACH_EMPTY As ARES_MS_VAR_Class 'Default Value: False                     True or False (Boolean)
' Used in PropertyRendering module — master switch for the text-render engine: a "Prop[Name]" token typed in a text is FULLY REPLACED by that custom property's value, the binding living in hidden ARES_SYS/ARES_Render metadata.
Public ARES_TEXT_RENDER As ARES_MS_VAR_Class 'Default Value: False                     True or False (Boolean)
' Used in PropertyActuator module — master switch for the Color-actuation engine (SELF reaction: reflects the fixed, reserved ARES_Color property's value onto the element's own Color; pilot property is not configurable, see ARESConstants.ARES_PROP_COLOR).
Public ARES_ACTUATE_COLOR As ARES_MS_VAR_Class 'Default Value: False                     True or False (Boolean)
' Used in PropertyActuator module — master switch for the Level-actuation engine (SELF reaction: reflects the fixed, reserved ARES_Lvl property's value onto the element's own Level; pilot property is not configurable, see ARESConstants.ARES_PROP_LVL).
Public ARES_ACTUATE_LEVEL As ARES_MS_VAR_Class 'Default Value: False                     True or False (Boolean)
' Used in FormPlacement module — remembers each option form's last on-screen position. INTERNAL, managed automatically (do not edit by hand).
Public ARES_FORM_LAYOUT As ARES_MS_VAR_Class 'Default Value: (empty)   <FormName>=<xPct>:<yPct> entries delimited by | (ARES_VAR_DELIMITER)
' Used in ElementChangeHandler module — when True, ProcessElement runs synchronously inside
' IChangeTrackEvents_ElementChanged (Add/Modify/Delete-linked) instead of being queued to the next MicroStation
' idle event. Default True = standard behaviour. False falls back to the deferred (queued) path - see
' ElementChangeHandler.cls header for the reentrancy guard and when to use the fallback.
Public ARES_DIRECT_PROCESSING As ARES_MS_VAR_Class 'Default Value: True                      True or False (Boolean)
' Used in GroupSplitGuard module — when True, a cut that leaves two or more cable geometries in one graphic group
' carrying ARES properties asks the user to separate the group, undo the cut, or keep it.
Public ARES_GROUP_SPLIT_GUARD As ARES_MS_VAR_Class 'Default Value: True                      True or False (Boolean)
' === Cable Report (per-cable trenching length by soil type) — own option set ===
' Used in CableReport module for the source level(s) of the cables to report on, aerial and underground, any cable element type
Public ARES_CABLEREPORT_CABLE_LEVEL As ARES_MS_VAR_Class 'Default Value: (empty)   |-delimited level list (ARES_VAR_DELIMITER); at least one must exist (required)
' Used in CableReport module for the trenching-zone level (RunOutline output); empty falls back to ARES_Outline_Output_Level
Public ARES_CABLEREPORT_ZONE_LEVEL As ARES_MS_VAR_Class 'Default Value: (empty)   Level name string
' Used in CableReport module — custom-property NAME read off each end cell (must be an ItemType of the "ARES" DGNLib)
Public ARES_CABLEREPORT_REPERE_PROPERTY As ARES_MS_VAR_Class 'Default Value: Repere   custom-property name
' Used in CableReport module — custom-property NAME of an underground cable's nature/type, read on the cable, then its graphic group
Public ARES_CABLEREPORT_NATURE_PROPERTY As ARES_MS_VAR_Class 'Default Value: Nature   custom-property name
' Used in CableReport module — custom-property NAME marking an aerial cable (read like the Nature one): aerial rows go on their own
' sheet, never measured against a trenching zone
Public ARES_CABLEREPORT_AERIAL_NATURE_PROPERTY As ARES_MS_VAR_Class 'Default Value: (empty)   custom-property name; empty = no aerial sheet
' Used in CableReport module — custom-property NAME read off the cable's linked text for its total length
Public ARES_CABLEREPORT_LONGUEUR_PROPERTY As ARES_MS_VAR_Class 'Default Value: Longueur   custom-property name
' Used in CableReport module — custom-property NAME read off each trenching zone; pivoted into one column per distinct value
Public ARES_CABLEREPORT_ZONE_PROPERTY As ARES_MS_VAR_Class 'Default Value: Coupe_Type   custom-property name
' Used in CableReport module — search radius (master units) for the nearest end-cell lookup around each cable endpoint
Public ARES_CABLEREPORT_SEARCH_RADIUS As ARES_MS_VAR_Class 'Default Value: 2.0   Search radius in master units (Double)
' Used in CableReport module — decimal places applied to the pivoted soil-type length sums
' Multiplies ARES_Outline_Distance to give the length below which a cable is CROSSING a zone rather
' than occupying it. Two trenches that intersect leave each other's cable inside for width/sin(angle),
' and the width IS 2 x that distance - so a factor of 2 is exactly one trench width, a PERPENDICULAR
' crossing's own value, and the test is strictly greater so that such a crossing is excluded exactly.
' The default 3.5 leaves room for oblique ones, which measure width/sin(angle) and so overshoot a
' single width: 0.5 m was measured at ~53 deg with a 0.2 distance, and 3.5 puts the threshold at
' 0.7 m. Raise it further on a site whose cables cross at very shallow angles; 0 disables the filter.
Public ARES_CABLEREPORT_CROSS_FACTOR As ARES_MS_VAR_Class 'Default Value: 3.5   Multiplier on ARES_Outline_Distance (Double)
Public ARES_CABLEREPORT_ROUND As ARES_MS_VAR_Class 'Default Value: 2   Range 0 to 254 (Byte -1) #255 it reserved for error (ARES_RND_ERROR_VALUE)
' Used in CableReport module — open the workbook in Excel after export when True; False = headless silent export
Public ARES_CABLEREPORT_EXCEL_VISIBLE As ARES_MS_VAR_Class 'Default Value: false   True or False (Boolean)
' === Sheet Levels (turn every level on in the "Folio" sheet models) ===
' Used in SheetLevels module — the NAME pattern(s) selecting which Sheet ("Papier") models
' ActivateSheetLevels acts on. Matched with RuleGrammar.LikeAnyInListCI: case-insensitive, */? wildcards,
' WHOLE-name match — so "contains Folio" is spelled "*Folio*", which is the shipped default. Several
' alternatives may be | -separated (each one trimmed). "[" / "]" are NOT supported in a pattern: Like
' reads them as a character class. Empty, or nothing but delimiters = the key-in does nothing (fail-closed).
Public ARES_SHEET_LEVELS_MODEL_NAME As ARES_MS_VAR_Class 'Default Value: *Folio*   |-separated wildcard patterns (ARES_VAR_DELIMITER)
' Used in SheetLevels module — when True (the default), ActivateSheetLevels also turns on the levels of
' every reference attached to a matching sheet, walked depth-first. False = that sheet's own levels only,
' which leaves a folio whose content comes from a reference blank. Anything other than "False" reads as
' True: a half-done job is the worse failure here.
Public ARES_SHEET_LEVELS_ATTACHMENTS As ARES_MS_VAR_Class 'Default Value: True          True or False (Boolean)
' Used in ConfigThemes module — the theme last loaded or saved through the theme window; ImportARESConfig clears it
Public ARES_THEME_CURRENT As ARES_MS_VAR_Class 'Default Value: (empty)   theme name, without .cfg

' Initialize the configuration manager
Public Function Initialize() As Boolean
    On Error GoTo ErrorHandler

    Initialize = False
    Set moConfigVariables = New Collection
    mbInitialized = False

    ' Create all configuration variables
    If Not CreateConfigVariables() Then Exit Function

    mbInitialized = True
    Initialize = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfig.Initialize"
    Initialize = False
End Function

' Create all configuration variables
Private Function CreateConfigVariables() As Boolean
    On Error GoTo ErrorHandler

    CreateConfigVariables = False

    ' Create each configuration variable. Station = this machine's (never in a theme); business = what a theme carries.
    Set ARES_LANGUAGE = CreateConfigVar("ARES_Language", "", ARESVarScopeStation)
    Set ARES_UPDATE_ATLASCELLLABEL = CreateConfigVar("ARES_Update_ATLASCellLabel", "True", ARESVarScopeBusiness)
    Set ARES_CELL_LIKE_LABEL = CreateConfigVar("ARES_Cell_Is_Label_Name", "ETI076|ETI01E|ETI01F|ETI01K|ETI01M|ETI01N|ETI01O", ARESVarScopeBusiness)
    Set ARES_BULK_THRESHOLD = CreateConfigVar("ARES_Bulk_Threshold", "10", ARESVarScopeStation)
    Set ARES_BULK_INTERVAL = CreateConfigVar("ARES_Bulk_Interval", "1000", ARESVarScopeStation)
    Set ARES_UPDATE_MUTE = CreateConfigVar("ARES_Update_Mute", "False", ARESVarScopeStation)
    Set ARES_UPDATE_IGNORE_VERSION = CreateConfigVar("ARES_Update_Ignore_Version", "", ARESVarScopeStation)
    Set ARES_ZONING_LEVEL = CreateConfigVar("ARES_Zoning_Level", "", ARESVarScopeBusiness)
    Set ARES_ZONING_OUTPUT_LEVEL = CreateConfigVar("ARES_Zoning_Output_Level", "ARES_Zoning", ARESVarScopeBusiness)
    Set ARES_ZONING_OUTPUT_COLOR = CreateConfigVar("ARES_Zoning_Output_Color", "5", ARESVarScopeBusiness)
    Set ARES_ZONING_OUTPUT_STYLE = CreateConfigVar("ARES_Zoning_Output_Style", "0", ARESVarScopeBusiness)
    Set ARES_ZONING_OUTPUT_WEIGHT = CreateConfigVar("ARES_Zoning_Output_Weight", "4", ARESVarScopeBusiness)
    Set ARES_ZONING_DISTANCE = CreateConfigVar("ARES_Zoning_Distance", "2.0", ARESVarScopeBusiness)
    Set ARES_OUTLINE_LEVEL = CreateConfigVar("ARES_Outline_Level", "", ARESVarScopeBusiness)
    Set ARES_OUTLINE_DISTANCE = CreateConfigVar("ARES_Outline_Distance", "0.2", ARESVarScopeBusiness)
    Set ARES_OUTLINE_OUTPUT_LEVEL = CreateConfigVar("ARES_Outline_Output_Level", "ARES_Outline", ARESVarScopeBusiness)
    Set ARES_OUTLINE_OUTPUT_COLOR = CreateConfigVar("ARES_Outline_Output_Color", "5", ARESVarScopeBusiness)
    Set ARES_OUTLINE_OUTPUT_STYLE = CreateConfigVar("ARES_Outline_Output_Style", "0", ARESVarScopeBusiness)
    Set ARES_OUTLINE_OUTPUT_WEIGHT = CreateConfigVar("ARES_Outline_Output_Weight", "4", ARESVarScopeBusiness)
    Set ARES_ZONE_EXPORT_EXCEL_VISIBLE = CreateConfigVar("ARES_Zone_Export_Excel_Visible", "False", ARESVarScopeBusiness)
    Set ARES_ZONE_EXPORT_SHEET_NAME = CreateConfigVar("ARES_Zone_Export_Sheet_Name", "Export", ARESVarScopeBusiness)
    Set ARES_ZONE_EXPORT_USE_DIALOG = CreateConfigVar("ARES_Zone_Export_Use_Dialog", "True", ARESVarScopeBusiness)
    Set ARES_ZONE_EXPORT_ROUND = CreateConfigVar("ARES_Zone_Export_Round", "2", ARESVarScopeBusiness)
    Set ARES_ZONE_EXPORT_GROUP_BY = CreateConfigVar("ARES_Zone_Export_Group_By", "Style", ARESVarScopeBusiness)
    Set ARES_ZONE_EXPORT_LEVEL = CreateConfigVar("ARES_Zone_Export_Level", "", ARESVarScopeBusiness)
    Set ARES_ZONE_EXPORT_PER_ZONE = CreateConfigVar("ARES_Zone_Export_Per_Zone", "False", ARESVarScopeBusiness)
    Set ARES_ZONE_EXPORT_ZONE_PROPERTY = CreateConfigVar("ARES_Zone_Export_Zone_Property", "", ARESVarScopeBusiness)
    Set ARES_REGIONSPLIT_COLLINEAR_TOL = CreateConfigVar("ARES_RegionSplit_Collinear_Tol", "0.000001", ARESVarScopeBusiness)
    Set ARES_REGIONSPLIT_KEEP_ORIGINAL = CreateConfigVar("ARES_RegionSplit_Keep_Original", "False", ARESVarScopeBusiness)
    Set ARES_REGIONSPLIT_STROKE_TOL = CreateConfigVar("ARES_RegionSplit_Stroke_Tol", "0.01", ARESVarScopeBusiness)
    Set ARES_REGIONMERGE_KEEP_ORIGINALS = CreateConfigVar("ARES_RegionMerge_Keep_Originals", "False", ARESVarScopeBusiness)
    Set ARES_AUTO_PROPERTIES = CreateConfigVar("ARES_Auto_Properties", "True", ARESVarScopeBusiness)
    Set ARES_PROPERTY_RULES = CreateConfigVar("ARES_Property_Rules", "", ARESVarScopeBusiness)
    Set ARES_CALC_RULES = CreateConfigVar("ARES_Calc_Rules", "", ARESVarScopeBusiness)
    Set ARES_PROPERTY_CALC = CreateConfigVar("ARES_Property_Calc", "True", ARESVarScopeBusiness)
    Set ARES_CALC_DETACH_EMPTY = CreateConfigVar("ARES_Calc_Detach_Empty", "False", ARESVarScopeBusiness)
    Set ARES_TEXT_RENDER = CreateConfigVar("ARES_Text_Render", "True", ARESVarScopeBusiness)
    Set ARES_ACTUATE_COLOR = CreateConfigVar("ARES_Actuate_Color", "False", ARESVarScopeBusiness)
    Set ARES_ACTUATE_LEVEL = CreateConfigVar("ARES_Actuate_Level", "False", ARESVarScopeBusiness)
    Set ARES_FORM_LAYOUT = CreateConfigVar("ARES_Form_Layout", "", ARESVarScopeStation)
    Set ARES_DIRECT_PROCESSING = CreateConfigVar("ARES_Direct_Processing", "True", ARESVarScopeStation)
    Set ARES_GROUP_SPLIT_GUARD = CreateConfigVar("ARES_Group_Split_Guard", "True", ARESVarScopeBusiness)
    Set ARES_CABLEREPORT_CABLE_LEVEL = CreateConfigVar("ARES_CableReport_Cable_Level", "", ARESVarScopeBusiness)
    Set ARES_CABLEREPORT_ZONE_LEVEL = CreateConfigVar("ARES_CableReport_Zone_Level", "", ARESVarScopeBusiness)
    Set ARES_CABLEREPORT_REPERE_PROPERTY = CreateConfigVar("ARES_CableReport_Repere_Property", "Repere", ARESVarScopeBusiness)
    Set ARES_CABLEREPORT_NATURE_PROPERTY = CreateConfigVar("ARES_CableReport_Nature_Property", "Nature", ARESVarScopeBusiness)
    Set ARES_CABLEREPORT_AERIAL_NATURE_PROPERTY = CreateConfigVar("ARES_CableReport_Aerial_Nature_Property", "", ARESVarScopeBusiness)
    Set ARES_CABLEREPORT_LONGUEUR_PROPERTY = CreateConfigVar("ARES_CableReport_Longueur_Property", "Longueur", ARESVarScopeBusiness)
    Set ARES_CABLEREPORT_ZONE_PROPERTY = CreateConfigVar("ARES_CableReport_Zone_Property", "Coupe_Type", ARESVarScopeBusiness)
    Set ARES_CABLEREPORT_SEARCH_RADIUS = CreateConfigVar("ARES_CableReport_Search_Radius", "2.0", ARESVarScopeBusiness)
    Set ARES_CABLEREPORT_CROSS_FACTOR = CreateConfigVar("ARES_CableReport_Cross_Factor", "3.5", ARESVarScopeBusiness)
    Set ARES_CABLEREPORT_ROUND = CreateConfigVar("ARES_CableReport_Round", "2", ARESVarScopeBusiness)
    Set ARES_CABLEREPORT_EXCEL_VISIBLE = CreateConfigVar("ARES_CableReport_Excel_Visible", "False", ARESVarScopeBusiness)
    Set ARES_SHEET_LEVELS_MODEL_NAME = CreateConfigVar("ARES_Sheet_Levels_Model_Name", "*Folio*", ARESVarScopeBusiness)
    Set ARES_SHEET_LEVELS_ATTACHMENTS = CreateConfigVar("ARES_Sheet_Levels_Attachments", "True", ARESVarScopeBusiness)
    Set ARES_THEME_CURRENT = CreateConfigVar("ARES_Theme_Current", "", ARESVarScopeStation)
    CreateConfigVariables = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfig.CreateConfigVariables"
    CreateConfigVariables = False
End Function

' Create a single configuration variable. Raises on a missing scope, so a new variable cannot fall on either
' side of a theme by default (CreateConfigVariables then fails and logs it).
Private Function CreateConfigVar(ByVal Key As String, ByVal DefaultValue As String, ByVal Scope As ARESVarScope) As ARES_MS_VAR_Class
    If Scope <> ARESVarScopeStation And Scope <> ARESVarScopeBusiness Then
        Err.Raise vbObjectError + 1, "ARESConfigClass.CreateConfigVar", "No valid scope declared for " & Key
    End If

    Dim ConfigVar As ARES_MS_VAR_Class
    Set ConfigVar = New ARES_MS_VAR_Class

    ConfigVar.Key = Key
    ConfigVar.DefaultValue = DefaultValue
    ConfigVar.Scope = Scope

    ' Get current value from MicroStation
    Dim currentValue As String
    currentValue = Config.GetVar(Key)

    ' If variable is not defined, set to default
    If currentValue = ARES_NAVD Then
        currentValue = DefaultValue
        Config.SetVar Key, DefaultValue
    End If

    ConfigVar.Value = currentValue
    moConfigVariables.Add ConfigVar, Key
    Set CreateConfigVar = ConfigVar
End Function

' Get a configuration variable by key
Public Function GetConfigVar(ByVal Key As String) As ARES_MS_VAR_Class
    On Error Resume Next
    Set GetConfigVar = moConfigVariables(Key)
    If Err.Number <> 0 Then
        Set GetConfigVar = Nothing
    End If
    On Error GoTo 0
End Function

' Every variable key, in declaration order (0-based; a single "" when nothing is declared).
Public Function GetConfigKeys() As String()
    On Error GoTo ErrorHandler

    Dim sKeys() As String
    Dim ConfigVar As Variant
    Dim i As Long

    ReDim sKeys(0 To 0)
    If Not moConfigVariables Is Nothing Then
        If moConfigVariables.Count > 0 Then
            ReDim sKeys(0 To moConfigVariables.Count - 1)
            For Each ConfigVar In moConfigVariables
                sKeys(i) = ConfigVar.Key
                i = i + 1
            Next ConfigVar
        End If
    End If
    GetConfigKeys = sKeys
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfigClass.GetConfigKeys"
    ReDim sKeys(0 To 0)
    GetConfigKeys = sKeys
End Function

' Reset a configuration variable to its default value
Public Function ResetConfigVar(ByVal Key As String) As Boolean
    On Error GoTo ErrorHandler

    Dim ConfigVar As ARES_MS_VAR_Class
    Set ConfigVar = moConfigVariables(Key)

    ' Reset in MicroStation
    If Not Config.SetVar(Key, ConfigVar.DefaultValue) Then
        ResetConfigVar = False
        Exit Function
    End If

    ' Update in memory
    ConfigVar.ResetToDefault
    ResetConfigVar = True

    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfig.ResetConfigVar"
    ResetConfigVar = False
End Function

' Remove a configuration variable
Public Function RemoveConfigVar(ByVal Key As String, Optional showConfirmation As Boolean = True) As Boolean
    On Error GoTo ErrorHandler

    Dim ConfigVar As ARES_MS_VAR_Class
    Set ConfigVar = moConfigVariables(Key)

    ' Ask for confirmation
    If showConfirmation Then
        If MsgBox(GetTranslation("VarRemoveConfirm", Key), vbYesNo) = vbNo Then
            RemoveConfigVar = False
            Exit Function
        End If
    End If

    ' Remove from MicroStation
    If Not Config.RemoveValue(Key) Then
        RemoveConfigVar = False
        Exit Function
    End If

    ' Remove from collection
    moConfigVariables.Remove Key
    RemoveConfigVar = True

    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfig.RemoveConfigVar"
    RemoveConfigVar = False
End Function

' Reset all configuration variables to their default values
Public Function ResetAllConfigVars() As Boolean
    On Error GoTo ErrorHandler

    Dim var As Variant
    Dim Success As Boolean

    Success = True
    For Each var In moConfigVariables
        If Not ResetConfigVar(var.Key) Then
            Success = False
        End If
    Next var

    ResetAllConfigVars = Success
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfig.ResetAllConfigVars"
    ResetAllConfigVars = False
End Function

' Remove all configuration variables
Public Function RemoveAllConfigVars() As Boolean
    On Error GoTo ErrorHandler

    Dim var As Variant
    Dim Success As Boolean

    ' Ask for confirmation
    If MsgBox(GetTranslation("VarsRemoveConfirm"), vbYesNo) = vbNo Then
        RemoveAllConfigVars = False
        Exit Function
    End If

    ' Collect keys first to avoid modifying the collection during For Each iteration
    Dim keysToRemove() As String
    Dim keyCount As Long
    Dim k As Long
    keyCount = moConfigVariables.Count
    ReDim keysToRemove(1 To keyCount)
    k = 1
    For Each var In moConfigVariables
        keysToRemove(k) = var.Key
        k = k + 1
    Next var

    Success = True
    For k = 1 To keyCount
        If Not RemoveConfigVar(keysToRemove(k), False) Then
            Success = False
        End If
    Next k

    RemoveAllConfigVars = Success
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfig.RemoveAllConfigVars"
    RemoveAllConfigVars = False
End Function

' Check if the configuration has been initialized
Public Property Get IsInitialized() As Boolean
    IsInitialized = mbInitialized
End Property

' Export configuration to file. bBusinessOnly: only the variables a theme carries (a theme save).
Public Function ExportConfig(Optional ByVal Filepath As String = "", Optional ByVal bBusinessOnly As Boolean = False) As Boolean
    On Error GoTo ErrorHandler

    ExportConfig = False

    If Not mbInitialized Then
        ErrorHandler.HandleError "Configuration not initialized", 0, "", "ARESConfigClass.ExportConfig"
        Exit Function
    End If

    ' Generate default file path if not provided
    If Len(Filepath) = 0 Then
        Filepath = GetDefaultExportPath()
    End If

    Dim FileNum As Integer
    Dim ConfigVar As Variant
    Dim ExportData As String

    ' Build export header
    ExportData = "# ARES Configuration Export" & vbCrLf
    ExportData = ExportData & "# Generated: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    ExportData = ExportData & "# Version: " & GetARESVersion() & vbCrLf
    ExportData = ExportData & vbCrLf

    ' Export each configuration variable (business only: a variable of no known scope is left out)
    For Each ConfigVar In moConfigVariables
        If Not bBusinessOnly Then
            ExportData = ExportData & FormatConfigForExport(ConfigVar) & vbCrLf
        ElseIf ConfigVar.Scope = ARESVarScopeBusiness Then
            ExportData = ExportData & FormatConfigForExport(ConfigVar) & vbCrLf
        End If
    Next ConfigVar
    
    ' Write to file
    FileNum = FreeFile
    Open Filepath For Output As #FileNum
    Print #FileNum, ExportData
    Close #FileNum
    
    ExportConfig = True
    Exit Function
    
ErrorHandler:
    If FileNum > 0 Then Close #FileNum
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfigClass.ExportConfig"
    ExportConfig = False
End Function

' Import configuration from file, every scope. The unknown-line count and the file's "# Version:" header come
' back through nUnknown / sFileVersion for the caller to report; nothing is logged for them.
Public Function ImportConfig(ByVal Filepath As String, Optional ByVal OverwriteExisting As Boolean = True, _
                             Optional ByRef nUnknown As Long, Optional ByRef sFileVersion As String) As Boolean
    On Error GoTo ErrorHandler

    ImportConfig = False

    If Not mbInitialized Then
        ErrorHandler.HandleError "Configuration not initialized", 0, "", "ARESConfigClass.ImportConfig"
        Exit Function
    End If

    Dim bReadable As Boolean
    ImportConfig = LoadConfig(Filepath, False, OverwriteExisting, bReadable, nUnknown, sFileVersion)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfigClass.ImportConfig"
    ImportConfig = False
End Function

' Fail-closed load: the whole file is read before the first variable is written. Writes nothing and returns
' False when the file cannot be read (bReadable False) or holds no recognised line of the scope asked for.
' bBusinessOnly: a theme load - station lines are ignored without counting as unknown.
Public Function LoadConfig(ByVal Filepath As String, ByVal bBusinessOnly As Boolean, ByVal OverwriteExisting As Boolean, _
                           ByRef bReadable As Boolean, ByRef nUnknown As Long, ByRef sFileVersion As String) As Boolean
    On Error GoTo ErrorHandler

    Dim sKeys() As String
    Dim sValues() As String
    Dim nLines As Long
    Dim nStation As Long
    Dim nBusiness As Long

    LoadConfig = False
    bReadable = ReadConfigFile(Filepath, sKeys, sValues, nLines, nStation, nBusiness, nUnknown, sFileVersion)
    If Not bReadable Then Exit Function

    If bBusinessOnly Then
        If nBusiness = 0 Then Exit Function
    ElseIf nLines = 0 Then
        Exit Function
    End If

    ApplyConfigLines sKeys, sValues, nLines, bBusinessOnly, OverwriteExisting
    LoadConfig = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfigClass.LoadConfig"
    LoadConfig = False
End Function

' The parser seam of every load: reads a whole config file and writes nothing. sKeys/sValues get the recognised
' lines in file order (0-based, nLines of them), nStation/nBusiness count them per scope, nUnknown counts the
' other non-comment lines, sFileVersion is the "# Version:" header ("" when absent). False = the file could not
' be read (missing, locked): expected, so not logged - the caller reports it.
Public Function ReadConfigFile(ByVal Filepath As String, ByRef sKeys() As String, ByRef sValues() As String, _
                               ByRef nLines As Long, ByRef nStation As Long, ByRef nBusiness As Long, _
                               ByRef nUnknown As Long, ByRef sFileVersion As String) As Boolean
    On Error GoTo ErrorHandler

    Dim FileNum As Integer
    Dim sLine As String
    Dim sTrimmed As String
    Dim sKey As String
    Dim sValue As String
    Dim ConfigVar As ARES_MS_VAR_Class
    Dim nErr As Long
    Dim sErrDesc As String
    Dim sErrSrc As String

    ReadConfigFile = False
    nLines = 0
    nStation = 0
    nBusiness = 0
    nUnknown = 0
    sFileVersion = ""
    ReDim sKeys(0 To 0)
    ReDim sValues(0 To 0)
    If Not mbInitialized Then Exit Function

    FileNum = FreeFile
    Open Filepath For Input As #FileNum
    Do While Not EOF(FileNum)
        Line Input #FileNum, sLine
        sTrimmed = Trim(sLine)

        If Left(sTrimmed, 10) = "# Version:" Then
            sFileVersion = Trim(Mid(sTrimmed, 11))
        ElseIf Len(sTrimmed) > 0 And Left(sTrimmed, 1) <> "#" Then
            Set ConfigVar = Nothing
            If ParseConfigLine(sLine, sKey, sValue) Then Set ConfigVar = GetConfigVar(sKey)

            If ConfigVar Is Nothing Then
                nUnknown = nUnknown + 1
            Else
                ReDim Preserve sKeys(0 To nLines)
                ReDim Preserve sValues(0 To nLines)
                sKeys(nLines) = ConfigVar.Key
                sValues(nLines) = sValue
                nLines = nLines + 1
                If ConfigVar.Scope = ARESVarScopeStation Then
                    nStation = nStation + 1
                ElseIf ConfigVar.Scope = ARESVarScopeBusiness Then
                    nBusiness = nBusiness + 1
                End If
            End If
        End If
    Loop
    Close #FileNum
    FileNum = 0

    ReadConfigFile = True
    Exit Function

ErrorHandler:
    nErr = Err.Number
    sErrDesc = Err.Description
    sErrSrc = Err.Source
    If FileNum > 0 Then Close #FileNum
    ReadConfigFile = False
    If Not IsFileAccessError(nErr) Then
        ErrorHandler.HandleError sErrDesc, nErr, sErrSrc, "ARESConfigClass.ReadConfigFile"
    End If
End Function

' Get configuration as formatted string (for preview)
Public Function GetConfigSummary() As String
    On Error GoTo ErrorHandler
    
    Dim Summary As String
    Dim ConfigVar As Variant
    
    Summary = "=== ARES Configuration Summary ===" & vbCrLf & vbCrLf
    
    For Each ConfigVar In moConfigVariables
        Summary = Summary & ConfigVar.Key & " = " & ConfigVar.Value
        If Not ConfigVar.IsDefault() Then
            Summary = Summary & " (Modified from: " & ConfigVar.DefaultValue & ")"
        End If
        Summary = Summary & vbCrLf
    Next ConfigVar
    
    GetConfigSummary = Summary
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfigClass.GetConfigSummary"
    GetConfigSummary = "Error generating summary"
End Function

Public Function GetDefaultExportPath() As String
    On Error Resume Next
    
    Dim Path As String
    
    If Not ActiveDesignFile Is Nothing Then
        Path = ActiveDesignFile.Path & "\ARES_Config_" & Format(Now, "yyyymmdd_hhmmss") & ".cfg"
    Else
        Path = Environ("USERPROFILE") & "\Desktop\ARES_Config_" & Format(Now, "yyyymmdd_hhmmss") & ".cfg"
    End If
    
    GetDefaultExportPath = Path
End Function

' === PRIVATE HELPER METHODS ===

Private Function FormatConfigForExport(ByVal ConfigVar As ARES_MS_VAR_Class) As String
    Dim Result As String

    ' Format: KEY=VALUE:::DEFAULT=defaultvalue:::MODIFIED=true/false
    ' ":::" is used instead of "|" to avoid conflict with ARES_VAR_DELIMITER used inside config values
    Result = ConfigVar.Key & "=" & ConfigVar.Value & ":::DEFAULT=" & ConfigVar.DefaultValue
    Result = Result & ":::MODIFIED=" & CStr(ConfigVar.IsModified())

    FormatConfigForExport = Result
End Function

' Split one line of the format KEY=VALUE:::DEFAULT=defaultvalue:::MODIFIED=true/false. False = no key.
Private Function ParseConfigLine(ByVal Line As String, ByRef ConfigKey As String, ByRef ConfigValue As String) As Boolean
    On Error GoTo ErrorHandler

    ParseConfigLine = False

    Dim Parts() As String
    Dim EqPos As Long

    Parts = Split(Line, ":::")
    If UBound(Parts) < 0 Then Exit Function

    ' Use InStr to find only the first "=" so values containing "=" are not truncated
    EqPos = InStr(Parts(0), "=")
    If EqPos < 2 Then Exit Function

    ConfigKey = Trim(Left(Parts(0), EqPos - 1))
    ConfigValue = Trim(Mid(Parts(0), EqPos + 1))
    ParseConfigLine = (Len(ConfigKey) > 0)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfigClass.ParseConfigLine"
    ParseConfigLine = False
End Function

' Write the lines ReadConfigFile returned: every scope, or business variables only. OverwriteExisting False
' keeps a variable already modified from its default.
Private Sub ApplyConfigLines(ByRef sKeys() As String, ByRef sValues() As String, ByVal nLines As Long, _
                             ByVal bBusinessOnly As Boolean, ByVal OverwriteExisting As Boolean)
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim ConfigVar As ARES_MS_VAR_Class

    For i = 0 To nLines - 1
        Set ConfigVar = GetConfigVar(sKeys(i))
        If ShouldApplyLine(ConfigVar, bBusinessOnly, OverwriteExisting) Then
            ' ConfigVar.Value Let persists to MicroStation via Config.SetVar
            ConfigVar.Value = sValues(i)
        End If
    Next i
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfigClass.ApplyConfigLines"
End Sub

Private Function ShouldApplyLine(ByVal ConfigVar As ARES_MS_VAR_Class, ByVal bBusinessOnly As Boolean, _
                                 ByVal OverwriteExisting As Boolean) As Boolean
    ShouldApplyLine = False
    If ConfigVar Is Nothing Then Exit Function
    If bBusinessOnly Then
        If ConfigVar.Scope <> ARESVarScopeBusiness Then Exit Function
    End If
    If Not OverwriteExisting Then
        If Not ConfigVar.IsDefault() Then Exit Function
    End If
    ShouldApplyLine = True
End Function

' Run-time errors of a file that cannot be opened or read: missing, locked, bad path, device not ready.
Private Function IsFileAccessError(ByVal nErr As Long) As Boolean
    Select Case nErr
        Case 52, 53, 54, 55, 57, 62, 67, 68, 70, 71, 74, 75, 76
            IsFileAccessError = True
        Case Else
            IsFileAccessError = False
    End Select
End Function

Private Function GetARESVersion() As String
    GetARESVersion = ARES_CONFIG_VERSION
End Function