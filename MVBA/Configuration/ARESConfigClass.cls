' Class Module: ARESConfigClass
' Description: Manages all ARES configuration variables
' License: This project is licensed under the AGPL-3.0.
' Dependencies: Config, ARES_MS_VAR_Class, ErrorHandlerClass
Option Explicit

Private moConfigVariables As Collection
Private mbInitialized As Boolean

'######################################################################################################################
'                             CAN BE MODIFIED IN MS ENVIRONMENT VARIABLES DO NOT MODIFY HERE
'                                   USE CONFIG MODULE TO GET, SET AND REMOVE A VALUE
'######################################################################################################################

' Used in Length module for default rounding
Public ARES_ROUNDS As ARES_MS_VAR_Class 'Default Value: 2  Range 0 to 254 (Byte -1)    #255 it reserved for error (ARES_RND_ERROR_VALUE)
' Used in ElementChangeHandler ClassModule for automatically enable adding length to a text if the conditions are met.
Public ARES_AUTO_LENGTHS As ARES_MS_VAR_Class 'Dafault Value: True          True or False (Boolean)
Public ARES_UPDATE_LENGTHS As ARES_MS_VAR_Class 'Dafault Value: True        True or False (Boolean)
' User in Auto_Lengths module for color updating
Public ARES_UPDATE_COLOR_WITH_LENGTH As ARES_MS_VAR_Class 'Dafault Value: True          True or False (Boolean)
' Used in Auto_Lengths module For length-specific rounding in text.
Public ARES_LENGTH_ROUND As ARES_MS_VAR_Class 'Default Value: 1             Range 0 to 254 (Byte -1) #255 it reserved for error (ARES_RND_ERROR_VALUE)
' Used in Auto_Lengths and StringsInEl module For Triggers in text.
Public ARES_LENGTH_TRIGGER As ARES_MS_VAR_Class 'Default Value: (Xx_m)      can a array use | (ARES_VAR_DELIMITER) like (Xx_m)|(Xx_cm)|(Xx_km)
' Used in AutoLengths and StringsInEl module for replace this triger with the length of element
Public ARES_LENGTH_TRIGGER_ID As ARES_MS_VAR_Class 'Default Value: Xx_      Must be present in triggers
' Used in LangManager module to force language if CONNECTUSER_LANGUAGE configuration variable is not set
Public ARES_LANGUAGE As ARES_MS_VAR_Class 'No Default Value
' Used in CellRedreaw module to chek if cell is like label (use cell name)
Public ARES_CELL_LIKE_LABEL As ARES_MS_VAR_Class 'Default Value:ETI076|ETI01E|ETI01F|ETI01K|ETI01M|ETI01N|ETI01O      can a array use | (ARES_VAR_DELIMITER)
' Used in CellRedreaw module for automatically enable cell update
Public ARES_UPDATE_ATLASCELLLABEL As ARES_MS_VAR_Class 'Dafault Value: True          True or False (Boolean)
' Used in UnitTesting module
Public ARES_UnitTesting As ARES_MS_VAR_Class 'Default Value: I'm a ARES test unit variable
' Used in ElementChangeHandler ClassModule for automatically enable only color update.
Public ARES_ONLY_COLOR As ARES_MS_VAR_Class 'Dafault Value: True        True or False (Boolean)
' Used in ElementChangeHandler for bulk operation detection threshold
Public ARES_BULK_THRESHOLD As ARES_MS_VAR_Class 'Default Value: 10    Number of calls to trigger bulk mode
' Used in ElementChangeHandler for bulk operation detection time window
Public ARES_BULK_INTERVAL As ARES_MS_VAR_Class 'Default Value: 1000   Time window in milliseconds
' Used in BootLoader/IdleEventHandler for periodic license re-validation throttle.
' Unit: seconds. Default: 3600 (1 hour). Valid range: 60..86400 (1 minute to 24 hours).
' Out-of-range values fall back to default with a logged WARNING.
Public ARES_LICENSE_RECHECK_INTERVAL As ARES_MS_VAR_Class 'Default Value: 3600   Range 60..86400 seconds
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
' Used in RunZoning2 command for default buffer distance
Public ARES_ZONING2_DISTANCE As ARES_MS_VAR_Class 'Default Value: 0.2             Buffer distance in master units (Double)
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
' Used in RegionSplit module — geometric epsilon: degenerate-segment / zero-parameter / coincident-point threshold; also drives the cut-knife half-width
Public ARES_REGIONSPLIT_COLLINEAR_TOL As ARES_MS_VAR_Class 'Default Value: 0.000001               Geometric epsilon in master units (Double)
' Used in RegionSplit module — False = original region is deleted after the two halves are written; True = original is kept alongside the two halves
Public ARES_REGIONSPLIT_KEEP_ORIGINAL As ARES_MS_VAR_Class 'Default Value: False                  True or False (Boolean)
' Used in RegionSplit module — max chordal deviation (master units) when densifying an arc boundary side into a polyline; smaller => more chords => closer-to-radial perpendicular and tighter entry foot
Public ARES_REGIONSPLIT_STROKE_TOL As ARES_MS_VAR_Class 'Default Value: 0.01                       Stroke chordal tolerance in master units (Double, > 0)
' Used in CustomPropertyHandler module — allowed values for the "Coupe Type" custom property (dropdown choices).
' NOT hard-coded in code: edit this variable to change the list. Array delimited by | (ARES_VAR_DELIMITER).
Public ARES_COUPE_TYPE_LIST As ARES_MS_VAR_Class 'Default Value: TR1C|TR1A|TR1D    can be an array use | (ARES_VAR_DELIMITER)

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

    ' Create each configuration variable
    Set ARES_ROUNDS = CreateConfigVar("ARES_Round", "2")
    Set ARES_AUTO_LENGTHS = CreateConfigVar("ARES_Auto_Lengths", "True")
    Set ARES_UPDATE_LENGTHS = CreateConfigVar("ARES_Update_Lengths", "True")
    Set ARES_LENGTH_ROUND = CreateConfigVar("ARES_Length_Round", "1")
    Set ARES_LENGTH_TRIGGER = CreateConfigVar("ARES_Length_Triggers", "(Xx_m)")
    Set ARES_LENGTH_TRIGGER_ID = CreateConfigVar("ARES_Length_Trigger_ID", "Xx_")
    Set ARES_LANGUAGE = CreateConfigVar("ARES_Language", "")
    Set ARES_UnitTesting = CreateConfigVar("ARES_Unit_testing", "I'm a ARES test unit variable")
    Set ARES_UPDATE_COLOR_WITH_LENGTH = CreateConfigVar("ARES_Update_Color_With_Length", "True")
    Set ARES_UPDATE_ATLASCELLLABEL = CreateConfigVar("ARES_Update_ATLASCellLabel", "True")
    Set ARES_CELL_LIKE_LABEL = CreateConfigVar("ARES_Cell_Is_Label_Name", "ETI076|ETI01E|ETI01F|ETI01K|ETI01M|ETI01N|ETI01O")
    Set ARES_ONLY_COLOR = CreateConfigVar("ARES_Only_Color_Update", "True")
    Set ARES_BULK_THRESHOLD = CreateConfigVar("ARES_Bulk_Threshold", "10")
    Set ARES_BULK_INTERVAL = CreateConfigVar("ARES_Bulk_Interval", "1000")
    Set ARES_LICENSE_RECHECK_INTERVAL = CreateConfigVar("ARES_License_Recheck_Interval", "3600")
    Set ARES_UPDATE_MUTE = CreateConfigVar("ARES_Update_Mute", "False")
    Set ARES_UPDATE_IGNORE_VERSION = CreateConfigVar("ARES_Update_Ignore_Version", "")
    Set ARES_ZONING_LEVEL = CreateConfigVar("ARES_Zoning_Level", "")
    Set ARES_ZONING_OUTPUT_LEVEL = CreateConfigVar("ARES_Zoning_Output_Level", "ARES_Zoning")
    Set ARES_ZONING_OUTPUT_COLOR = CreateConfigVar("ARES_Zoning_Output_Color", "5")
    Set ARES_ZONING_OUTPUT_STYLE = CreateConfigVar("ARES_Zoning_Output_Style", "0")
    Set ARES_ZONING_OUTPUT_WEIGHT = CreateConfigVar("ARES_Zoning_Output_Weight", "4")
    Set ARES_ZONING_DISTANCE = CreateConfigVar("ARES_Zoning_Distance", "2.0")
    Set ARES_ZONING2_DISTANCE = CreateConfigVar("ARES_Zoning2_Distance", "0.2")
    Set ARES_ZONE_EXPORT_EXCEL_VISIBLE = CreateConfigVar("ARES_Zone_Export_Excel_Visible", "False")
    Set ARES_ZONE_EXPORT_SHEET_NAME = CreateConfigVar("ARES_Zone_Export_Sheet_Name", "Export")
    Set ARES_ZONE_EXPORT_USE_DIALOG = CreateConfigVar("ARES_Zone_Export_Use_Dialog", "True")
    Set ARES_ZONE_EXPORT_ROUND = CreateConfigVar("ARES_Zone_Export_Round", "2")
    Set ARES_ZONE_EXPORT_GROUP_BY = CreateConfigVar("ARES_Zone_Export_Group_By", "Style")
    Set ARES_REGIONSPLIT_COLLINEAR_TOL = CreateConfigVar("ARES_RegionSplit_Collinear_Tol", "0.000001")
    Set ARES_REGIONSPLIT_KEEP_ORIGINAL = CreateConfigVar("ARES_RegionSplit_Keep_Original", "False")
    Set ARES_REGIONSPLIT_STROKE_TOL = CreateConfigVar("ARES_RegionSplit_Stroke_Tol", "0.01")
    Set ARES_COUPE_TYPE_LIST = CreateConfigVar("ARES_Coupe_Type_List", "TR1C|TR1A|TR1D")
    CreateConfigVariables = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfig.CreateConfigVariables"
    CreateConfigVariables = False
End Function

' Create a single configuration variable
Private Function CreateConfigVar(ByVal Key As String, ByVal DefaultValue As String) As ARES_MS_VAR_Class
    Dim ConfigVar As ARES_MS_VAR_Class
    Set ConfigVar = New ARES_MS_VAR_Class

    ConfigVar.Key = Key
    ConfigVar.DefaultValue = DefaultValue

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

' Export configuration to file
Public Function ExportConfig(Optional ByVal Filepath As String = "") As Boolean
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
    
    ' Export each configuration variable
    For Each ConfigVar In moConfigVariables
        ExportData = ExportData & FormatConfigForExport(ConfigVar) & vbCrLf
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

' Import configuration from file
Public Function ImportConfig(ByVal Filepath As String, Optional ByVal OverwriteExisting As Boolean = True) As Boolean
    On Error GoTo ErrorHandler
    
    ImportConfig = False
    
    If Not mbInitialized Then
        ErrorHandler.HandleError "Configuration not initialized", 0, "", "ARESConfigClass.ImportConfig"
        Exit Function
    End If
    
    If Len(dir(Filepath)) = 0 Then
        ErrorHandler.HandleError "Import file not found: " & Filepath, 0, "", "ARESConfigClass.ImportConfig"
        Exit Function
    End If
    
    Dim FileNum As Integer
    Dim Line As String
    Dim ImportedCount As Long
    Dim SkippedCount As Long
    
    ' Read and process file
    FileNum = FreeFile
    Open Filepath For Input As #FileNum
    
    Do While Not EOF(FileNum)
        Line Input #FileNum, Line
        
        ' Check version header
        If Left(Trim(Line), 10) = "# Version:" Then
            Dim fileVersion As String
            fileVersion = Trim(Mid(Trim(Line), 11))
            If fileVersion <> ARES_CONFIG_VERSION Then
                ErrorHandler.HandleError "Config version mismatch: file=" & fileVersion & ", expected=" & ARES_CONFIG_VERSION, 0, "", "ARESConfigClass.ImportConfig"
            End If
        End If

        ' Skip comments and empty lines
        If Len(Trim(Line)) > 0 And Left(Trim(Line), 1) <> "#" Then
            If ProcessImportLine(Line, OverwriteExisting) Then
                ImportedCount = ImportedCount + 1
            Else
                SkippedCount = SkippedCount + 1
            End If
        End If
    Loop
    
    Close #FileNum

    If SkippedCount > 0 Then
        ErrorHandler.HandleError SkippedCount & " line(s) skipped during import of: " & Filepath, 0, "", "ARESConfigClass.ImportConfig"
    End If

    ImportConfig = True
    Exit Function
    
ErrorHandler:
    If FileNum > 0 Then Close #FileNum
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfigClass.ImportConfig"
    ImportConfig = False
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

Private Function ProcessImportLine(ByVal Line As String, ByVal OverwriteExisting As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    ProcessImportLine = False
    
    ' Parse line format: KEY=VALUE:::DEFAULT=defaultvalue:::MODIFIED=true/false
    Dim Parts() As String
    Dim ConfigKey As String
    Dim ConfigValue As String
    Dim ConfigVar As ARES_MS_VAR_Class
    Dim EqPos As Long

    Parts = Split(Line, ":::")
    If UBound(Parts) < 0 Then Exit Function

    ' Use InStr to find only the first "=" so values containing "=" are not truncated
    EqPos = InStr(Parts(0), "=")
    If EqPos < 2 Then Exit Function

    ConfigKey = Trim(Left(Parts(0), EqPos - 1))
    ConfigValue = Trim(Mid(Parts(0), EqPos + 1))
    
    ' Check if key exists in our configuration
    Set ConfigVar = GetConfigVar(ConfigKey)
    If ConfigVar Is Nothing Then
        ErrorHandler.HandleError "Unknown configuration key: " & ConfigKey, 0, "", "ARESConfigClass.ProcessImportLine"
        Exit Function
    End If
    
    ' Check if we should overwrite
    If Not OverwriteExisting And Not ConfigVar.IsDefault() Then
        Exit Function ' Skip if already modified and not overwriting
    End If
    
    ' Update configuration (ConfigVar.Value Let persists to MicroStation via Config.SetVar)
    ConfigVar.Value = ConfigValue
    ProcessImportLine = True
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfigClass.ProcessImportLine"
    ProcessImportLine = False
End Function

Private Function GetARESVersion() As String
    GetARESVersion = ARES_CONFIG_VERSION
End Function