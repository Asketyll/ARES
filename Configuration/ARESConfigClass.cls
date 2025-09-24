' Class Module: ARESConfigClass
' Description: Manages all ARES configuration variables
' License: This project is licensed under the AGPL-3.0.
' Dependencies: Config, ARES_MS_VAR_Class, ErrorHandlerClass
Option Explicit

Private pConfigVariables As Collection
Private pInitialized As Boolean

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

' Initialize the configuration manager
Public Function Initialize() As Boolean
    On Error GoTo ErrorHandler

    Initialize = False
    Set pConfigVariables = New Collection
    pInitialized = False

    ' Create all configuration variables
    If Not CreateConfigVariables() Then Exit Function

    pInitialized = True
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
    
    CreateConfigVariables = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfig.CreateConfigVariables"
    CreateConfigVariables = False
End Function

' Create a single configuration variable
Private Function CreateConfigVar(ByVal key As String, ByVal defaultValue As String) As ARES_MS_VAR_Class
    Dim ConfigVar As ARES_MS_VAR_Class
    Set ConfigVar = New ARES_MS_VAR_Class

    ConfigVar.key = key
    ConfigVar.defaultValue = defaultValue

    ' Get current value from MicroStation
    Dim currentValue As String
    currentValue = Config.GetVar(key)

    ' If variable is not defined, set to default
    If currentValue = ARES_NAVD Then
        currentValue = defaultValue
        Config.SetVar key, defaultValue
    End If

    ConfigVar.Value = currentValue
    pConfigVariables.Add ConfigVar, key
    Set CreateConfigVar = ConfigVar
End Function

' Get a configuration variable by key
Public Function GetConfigVar(ByVal key As String) As ARES_MS_VAR_Class
    On Error Resume Next
    Set GetConfigVar = pConfigVariables(key)
    If Err.Number <> 0 Then
        Set GetConfigVar = Nothing
    End If
    On Error GoTo 0
End Function

' Reset a configuration variable to its default value
Public Function ResetConfigVar(ByVal key As String) As Boolean
    On Error GoTo ErrorHandler

    Dim ConfigVar As ARES_MS_VAR_Class
    Set ConfigVar = pConfigVariables(key)

    If ConfigVar Is Nothing Then
        ResetConfigVar = False
        Exit Function
    End If

    ' Reset in MicroStation
    If Not Config.SetVar(key, ConfigVar.defaultValue) Then
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
Public Function RemoveConfigVar(ByVal key As String, Optional showConfirmation As Boolean = True) As Boolean
    On Error GoTo ErrorHandler

    Dim ConfigVar As ARES_MS_VAR_Class
    Set ConfigVar = pConfigVariables(key)

    If ConfigVar Is Nothing Then
        RemoveConfigVar = False
        Exit Function
    End If

    ' Ask for confirmation
    If showConfirmation Then
        If MsgBox(GetTranslation("VarRemoveConfirm", key), vbYesNo) = vbNo Then
            RemoveConfigVar = False
            Exit Function
        End If
    End If

    ' Remove from MicroStation
    If Not Config.RemoveValue(key) Then
        RemoveConfigVar = False
        Exit Function
    End If

    ' Remove from collection
    pConfigVariables.Remove key
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
    For Each var In pConfigVariables
        If Not ResetConfigVar(var.key) Then
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

    Success = True
    For Each var In pConfigVariables
        If Not RemoveConfigVar(var.key, False) Then
            Success = False
        End If
    Next var

    RemoveAllConfigVars = Success
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfig.RemoveAllConfigVars"
    RemoveAllConfigVars = False
End Function

' Check if the configuration has been initialized
Public Property Get IsInitialized() As Boolean
    IsInitialized = pInitialized
End Property

' Export configuration to file
Public Function ExportConfig(Optional ByVal FilePath As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ExportConfig = False
    
    If Not pInitialized Then
        ErrorHandler.HandleError "Configuration not initialized", 0, "ARESConfigClass.ExportConfig"
        Exit Function
    End If
    
    ' Generate default file path if not provided
    If Len(FilePath) = 0 Then
        FilePath = GetDefaultExportPath()
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
    For Each ConfigVar In pConfigVariables
        ExportData = ExportData & FormatConfigForExport(ConfigVar) & vbCrLf
    Next ConfigVar
    
    ' Write to file
    FileNum = FreeFile
    Open FilePath For Output As #FileNum
    Print #FileNum, ExportData
    Close #FileNum
    
    ExportConfig = True
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfigClass.ExportConfig"
    ExportConfig = False
End Function

' Import configuration from file
Public Function ImportConfig(ByVal FilePath As String, Optional ByVal OverwriteExisting As Boolean = True) As Boolean
    On Error GoTo ErrorHandler
    
    ImportConfig = False
    
    If Not pInitialized Then
        ErrorHandler.HandleError "Configuration not initialized", 0, "ARESConfigClass.ImportConfig"
        Exit Function
    End If
    
    If Len(Dir(FilePath)) = 0 Then
        ErrorHandler.HandleError "Import file not found: " & FilePath, 0, "ARESConfigClass.ImportConfig"
        Exit Function
    End If
    
    Dim FileNum As Integer
    Dim Line As String
    Dim ImportedCount As Integer
    Dim SkippedCount As Integer
    
    ' Read and process file
    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    
    Do While Not EOF(FileNum)
        Line Input #FileNum, Line
        
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
    
    ' Show results
    Dim ResultMsg As String
    ResultMsg = "Import completed: " & ImportedCount & " imported, " & SkippedCount & " skipped"
    ErrorHandler.HandleError ResultMsg, 0, "ARESConfigClass.ImportConfig"
    
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
    
    For Each ConfigVar In pConfigVariables
        Summary = Summary & ConfigVar.key & " = " & ConfigVar.Value
        If Not ConfigVar.IsDefault() Then
            Summary = Summary & " (Modified from: " & ConfigVar.defaultValue & ")"
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
    
    ' Format: KEY=VALUE|DEFAULT=defaultvalue|MODIFIED=true/false
    Result = ConfigVar.key & "=" & ConfigVar.Value & "|DEFAULT=" & ConfigVar.defaultValue
    Result = Result & "|MODIFIED=" & CStr(ConfigVar.IsModified())
    
    FormatConfigForExport = Result
End Function

Private Function ProcessImportLine(ByVal Line As String, ByVal OverwriteExisting As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    ProcessImportLine = False
    
    ' Parse line format: KEY=VALUE|DEFAULT=defaultvalue|MODIFIED=true/false
    Dim Parts() As String
    Dim KeyValue() As String
    Dim ConfigKey As String
    Dim ConfigValue As String
    Dim ConfigVar As ARES_MS_VAR_Class
    
    Parts = Split(Line, "|")
    If UBound(Parts) < 0 Then Exit Function
    
    KeyValue = Split(Parts(0), "=")
    If UBound(KeyValue) < 1 Then Exit Function
    
    ConfigKey = Trim(KeyValue(0))
    ConfigValue = Trim(KeyValue(1))
    
    ' Check if key exists in our configuration
    Set ConfigVar = GetConfigVar(ConfigKey)
    If ConfigVar Is Nothing Then
        ErrorHandler.HandleError "Unknown configuration key: " & ConfigKey, 0, "ARESConfigClass.ProcessImportLine"
        Exit Function
    End If
    
    ' Check if we should overwrite
    If Not OverwriteExisting And Not ConfigVar.IsDefault() Then
        Exit Function ' Skip if already modified and not overwriting
    End If
    
    ' Update configuration
    If Config.SetVar(ConfigKey, ConfigValue) Then
        ConfigVar.Value = ConfigValue
        ProcessImportLine = True
    End If
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfigClass.ProcessImportLine"
    ProcessImportLine = False
End Function

Private Function GetARESVersion() As String
    GetARESVersion = "1.0.0" ' Update with your actual version
End Function