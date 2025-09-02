' Class Module: ARESConfigClass
' Description: Central manager for all ARES configuration variables with MicroStation integration
' License: This project is licensed under the AGPL-3.0.
' Dependencies: Config, ARES_MS_VAR_Class, ErrorHandlerClass, ARESConstants
Option Explicit

' === PRIVATE MEMBERS ===
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
' Used in CustomPropertyHandler module for default name of Library Type object
Public ARES_NAME_LIBRARY_TYPE As ARES_MS_VAR_Class 'Default Value: ARES
' Used in CustomPropertyHandler module for default name of Item Type object
Public ARES_NAME_ITEM_TYPE As ARES_MS_VAR_Class 'Default Value: ARESAutoLengthObject
' Used in LangManager module to force language if CONNECTUSER_LANGUAGE configuration variable is not set
Public ARES_LANGUAGE As ARES_MS_VAR_Class 'No Default Value
' Used in CellRedreaw module to chek if cell is like label (use cell name)
Public ARES_CELL_LIKE_LABEL As ARES_MS_VAR_Class 'Default Value:ETI076|ETI01E|ETI01F|ETI01K|ETI01M|ETI01N|ETI01O      can a array use | (ARES_VAR_DELIMITER)
' Used in CellRedreaw module for automatically enable cell update
Public ARES_UPDATE_ATLASCELLLABEL As ARES_MS_VAR_Class 'Dafault Value: True          True or False (Boolean)
' Used in UnitTesting module
Public ARES_UnitTesting As ARES_MS_VAR_Class 'Default Value: I'm a ARES test unit variable

' Initialize the configuration manager and all ARES variables
Public Function Initialize() As Boolean
    On Error GoTo ErrorHandler

    Initialize = False
    
    ' Clean up any existing configuration
    If Not moConfigVariables Is Nothing Then
        Set moConfigVariables = Nothing
    End If
    
    Set moConfigVariables = New Collection
    mbInitialized = False

    ' Create all configuration variables
    If Not CreateConfigVariables() Then
        ErrorHandler.HandleError "Failed to create configuration variables", 0, "ARESConfig.Initialize", "ERROR"
        Exit Function
    End If

    mbInitialized = True
    Initialize = True
    
    ' Log successful initialization
    ErrorHandler.HandleError "ARES configuration initialized successfully", 0, "ARESConfig.Initialize", "INFO"
    
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfig.Initialize"
    Initialize = False
End Function

' Create and initialize all ARES configuration variables
Private Function CreateConfigVariables() As Boolean
    On Error GoTo ErrorHandler

    CreateConfigVariables = False
    
    ' Validate collection is ready
    If moConfigVariables Is Nothing Then
        ErrorHandler.HandleError "Configuration collection not initialized", 0, "ARESConfig.CreateConfigVariables", "ERROR"
        Exit Function
    End If

    ' Create each configuration variable with proper error checking
    Set ARES_ROUNDS = CreateConfigVar("ARES_Round", "2")
    If ARES_ROUNDS Is Nothing Then Exit Function
    
    Set ARES_AUTO_LENGTHS = CreateConfigVar("ARES_Auto_Lengths", "True")
    If ARES_AUTO_LENGTHS Is Nothing Then Exit Function
    
    Set ARES_UPDATE_LENGTHS = CreateConfigVar("ARES_Update_Lengths", "True")
    If ARES_UPDATE_LENGTHS Is Nothing Then Exit Function
    
    Set ARES_LENGTH_ROUND = CreateConfigVar("ARES_Length_Round", "1")
    If ARES_LENGTH_ROUND Is Nothing Then Exit Function
    
    Set ARES_LENGTH_TRIGGER = CreateConfigVar("ARES_Length_Triggers", "(Xx_m)")
    If ARES_LENGTH_TRIGGER Is Nothing Then Exit Function
    
    Set ARES_LENGTH_TRIGGER_ID = CreateConfigVar("ARES_Length_Trigger_ID", "Xx_")
    If ARES_LENGTH_TRIGGER_ID Is Nothing Then Exit Function
    
    Set ARES_NAME_LIBRARY_TYPE = CreateConfigVar("ARES_Library_Type_Name", "ARES")
    If ARES_NAME_LIBRARY_TYPE Is Nothing Then Exit Function
    
    Set ARES_NAME_ITEM_TYPE = CreateConfigVar("ARES_Item_Type_Name", "ARESAutoLengthObject")
    If ARES_NAME_ITEM_TYPE Is Nothing Then Exit Function
    
    Set ARES_LANGUAGE = CreateConfigVar("ARES_Language", "")
    If ARES_LANGUAGE Is Nothing Then Exit Function
    
    Set ARES_UnitTesting = CreateConfigVar("ARES_Unit_testing", "I'm a ARES test unit variable")
    If ARES_UnitTesting Is Nothing Then Exit Function
    
    Set ARES_UPDATE_COLOR_WITH_LENGTH = CreateConfigVar("ARES_Update_Color_With_Length", "True")
    If ARES_UPDATE_COLOR_WITH_LENGTH Is Nothing Then Exit Function
    
    Set ARES_UPDATE_ATLASCELLLABEL = CreateConfigVar("ARES_Update_ATLASCellLabel", "True")
    If ARES_UPDATE_ATLASCELLLABEL Is Nothing Then Exit Function
    
    Set ARES_CELL_LIKE_LABEL = CreateConfigVar("ARES_Cell_Is_Label_Name", "ETI076|ETI01E|ETI01F|ETI01K|ETI01M|ETI01N|ETI01O")
    If ARES_CELL_LIKE_LABEL Is Nothing Then Exit Function
    
    CreateConfigVariables = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfig.CreateConfigVariables"
    CreateConfigVariables = False
End Function

' Create and initialize a single configuration variable
Private Function CreateConfigVar(ByVal StrKey As String, ByVal strDefaultValue As String) As ARES_MS_VAR_Class
    On Error GoTo ErrorHandler
    
    Dim oConfigVar As ARES_MS_VAR_Class
    Set oConfigVar = New ARES_MS_VAR_Class
    
    ' Initialize the variable with key and default
    oConfigVar.Initialize StrKey, strDefaultValue
    
    ' Get current value from MicroStation configuration
    Dim strCurrentValue As String
    strCurrentValue = Config.GetVar(StrKey)
    
    ' If variable is not defined in MicroStation, create it with default value
    If strCurrentValue = ARESConstants.ARES_NAVD Then
        strCurrentValue = strDefaultValue
        If Not Config.SetVar(StrKey, strDefaultValue) Then
            ' Log warning if we can't set the variable but continue
            ErrorHandler.HandleError "Failed to set default value for: " & StrKey, 0, "ARESConfig.CreateConfigVar", "WARNING"
        End If
    End If
    
    ' Set the current value
    oConfigVar.Value = strCurrentValue
    
    ' Add to collection
    moConfigVariables.Add oConfigVar, StrKey
    Set CreateConfigVar = oConfigVar
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfig.CreateConfigVar"
    Set CreateConfigVar = Nothing
End Function

' Get a configuration variable by key
Public Function GetConfigVar(ByVal StrKey As String) As ARES_MS_VAR_Class
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If Not mbInitialized Then
        ErrorHandler.HandleError "Configuration not initialized", 0, "ARESConfig.GetConfigVar", "ERROR"
        Set GetConfigVar = Nothing
        Exit Function
    End If
    
    If Len(Trim(StrKey)) = 0 Then
        ErrorHandler.HandleError "Empty key provided", 0, "ARESConfig.GetConfigVar", "ERROR"
        Set GetConfigVar = Nothing
        Exit Function
    End If
    
    Set GetConfigVar = moConfigVariables(StrKey)
    Exit Function
    
ErrorHandler:
    ' Key not found or other error
    Set GetConfigVar = Nothing
End Function

' Reset a configuration variable to its default value
Public Function ResetConfigVar(ByVal Key As String) As Boolean
    On Error GoTo ErrorHandler

    Dim configVar As ARES_MS_VAR_Class
    Set configVar = pConfigVariables(Key)

    If configVar Is Nothing Then
        ResetConfigVar = False
        Exit Function
    End If

    ' Reset in MicroStation
    If Not Config.SetVar(Key, configVar.Default) Then
        ResetConfigVar = False
        Exit Function
    End If

    ' Update in memory
    configVar.ResetToDefault
    ResetConfigVar = True

    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARESConfig.ResetConfigVar"
    ResetConfigVar = False
End Function

' Remove a configuration variable
Public Function RemoveConfigVar(ByVal Key As String, Optional showConfirmation As Boolean = True) As Boolean
    On Error GoTo ErrorHandler

    Dim configVar As ARES_MS_VAR_Class
    Set configVar = pConfigVariables(Key)

    If configVar Is Nothing Then
        RemoveConfigVar = False
        Exit Function
    End If

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
    pConfigVariables.Remove Key
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

    Success = True
    For Each var In pConfigVariables
        If Not RemoveConfigVar(var.Key, False) Then
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
    IsInitialized = mbInitialized
End Property

' Get count of configuration variables
Public Property Get Count() As Long
    On Error GoTo ErrorHandler
    
    If moConfigVariables Is Nothing Then
        Count = 0
    Else
        Count = moConfigVariables.Count
    End If
    Exit Property
    
ErrorHandler:
    Count = 0
End Property

' Clean up when class is destroyed
Private Sub Class_Terminate()
    On Error Resume Next
    
    If Not moConfigVariables Is Nothing Then
        Set moConfigVariables = Nothing
    End If
    
    ' Clean up all variable references
    Set ARES_ROUNDS = Nothing
    Set ARES_AUTO_LENGTHS = Nothing
    Set ARES_UPDATE_LENGTHS = Nothing
    Set ARES_LENGTH_ROUND = Nothing
    Set ARES_LENGTH_TRIGGER = Nothing
    Set ARES_LENGTH_TRIGGER_ID = Nothing
    Set ARES_NAME_LIBRARY_TYPE = Nothing
    Set ARES_NAME_ITEM_TYPE = Nothing
    Set ARES_LANGUAGE = Nothing
    Set ARES_UnitTesting = Nothing
    Set ARES_UPDATE_COLOR_WITH_LENGTH = Nothing
    Set ARES_UPDATE_ATLASCELLLABEL = Nothing
    Set ARES_CELL_LIKE_LABEL = Nothing
    
    mbInitialized = False
End Sub