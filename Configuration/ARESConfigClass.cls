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
    Set ARES_NAME_LIBRARY_TYPE = CreateConfigVar("ARES_Library_Type_Name", "ARES")
    Set ARES_NAME_ITEM_TYPE = CreateConfigVar("ARES_Item_Type_Name", "ARESAutoLengthObject")
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
    Dim configVar As ARES_MS_VAR_Class
    Set configVar = New ARES_MS_VAR_Class

    configVar.key = key
    configVar.defaultValue = defaultValue

    ' Get current value from MicroStation
    Dim currentValue As String
    currentValue = Config.GetVar(key)

    ' If variable is not defined, set to default
    If currentValue = ARES_NAVD Then
        currentValue = defaultValue
        Config.SetVar key, defaultValue
    End If

    configVar.Value = currentValue
    pConfigVariables.Add configVar, key
    Set CreateConfigVar = configVar
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

    Dim configVar As ARES_MS_VAR_Class
    Set configVar = pConfigVariables(key)

    If configVar Is Nothing Then
        ResetConfigVar = False
        Exit Function
    End If

    ' Reset in MicroStation
    If Not Config.SetVar(key, configVar.defaultValue) Then
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
Public Function RemoveConfigVar(ByVal key As String, Optional showConfirmation As Boolean = True) As Boolean
    On Error GoTo ErrorHandler

    Dim configVar As ARES_MS_VAR_Class
    Set configVar = pConfigVariables(key)

    If configVar Is Nothing Then
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