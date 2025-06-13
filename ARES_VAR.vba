' Module: ARES_VAR
' Description: This module provides configuration variables, absolute values, and explanations of each variable and where they are used.

'######################################################################################################################
'                       CAN'T BE EDITED IN MS VARIABLES ENVIRONMENT, YOU CAN MODIFY HERE
'######################################################################################################################

' Used in Link module and ElementChangeHandler ClassModule for check if a graphics group exists
Public Const ARES_DEFAULT_GRAPHIC_GROUP_ID As Long = 0 ' Constant for no graphic group

' Used in Link and MicroStationDefinition module for check if a MsdElementType is unknow or raise a error
Public Const ARES_MSDETYPE_ERROR As Long = 44 ' If you use Type 44, you can replace with another MsdElementType not used

' Used in StringsInEl and AutoLengths module for separate a list in a environment variables
Public Const ARES_VAR_DELIMITER As String = "|" ' Delimiter used in :ARES_LENGTH_TRIGGER

' Used in Config and ARES_VAR
Public Const ARES_NAVD As String = "NaVD" ' Constant for undefined MS configuration variables :Not a Variable Defined

' Used in Length
Public Const ARES_RND_ERROR_VALUE As Byte = 255 ' Constant for error in ARES_ROUNDS and ARES_LENGTH_ROUND

'######################################################################################################################
'                             CAN BE MODIFIED IN MS ENVIRONMENT VARIABLES DO NOT MODIFY HERE
'                                   USE CONFIG MODULE TO GET, SET AND REMOVE A VALUE
'######################################################################################################################

Public Type ARES_MS_VAR
    key As String
    Value As String
    Default As String
End Type

' Used in Length module for default rounding
Public ARES_ROUNDS As ARES_MS_VAR 'Default Value: 2  Range 0 to 254 (Byte -1)    #255 it reserved for error (ARES_RND_ERROR_VALUE)

' Used in ElementChangeHandler ClassModule for automatically enable adding length to a text if the conditions are met.
Public ARES_AUTO_LENGTHS As ARES_MS_VAR 'Dafault Value: True (ARES_AUTO_LENGTHS_DEFAULT)     True or False (Boolean)

' Used in Auto_Lengths module For length-specific rounding in text.
Public ARES_LENGTH_ROUND As ARES_MS_VAR 'Default Value: 1 (ARES_LENGTH_RND_DEFAULT) Range 0 to 254 (Byte -1) #255 it reserved for error (ARES_RND_ERROR_VALUE)

' Used in Auto_Lengths and StringsInEl module For Triggers in text.
Public ARES_LENGTH_TRIGGER As ARES_MS_VAR 'Default Value: (Xx_m) (ARES_LENGTH_TRIGGER_DEFAULT) can a array use | (ARES_VAR_DELIMITER) like (Xx_m)|(Xx_cm)|(Xx_km)

' Used in AutoLengths and StringsInEl module for replace this triger with the length of element
Public ARES_LENGTH_TRIGGER_ID As ARES_MS_VAR 'Default Value: Xx_  (ARES_LENGTH_TRIGGER_ID_DEFAULT)

' Used in CustomPropertyHandler module for default name of Library Type object
Public ARES_NAME_LIBRARY_TYPE As ARES_MS_VAR 'Default Value: ARES

' Used in CustomPropertyHandler module for default name of Item Type object
Public ARES_NAME_ITEM_TYPE As ARES_MS_VAR 'Default Value: ARESAutoLengthObject

Public Function InitMSVars()
    InitializeMSVar ARES_ROUNDS, "ARES_Round", "2"
    InitializeMSVar ARES_AUTO_LENGTHS, "ARES_Auto_Lengths", "True"
    InitializeMSVar ARES_LENGTH_ROUND, "ARES_Length_Round", "1"
    InitializeMSVar ARES_LENGTH_TRIGGER, "ARES_Length_Triggers", "(Xx_m)"
    InitializeMSVar ARES_LENGTH_TRIGGER_ID, "ARES_Length_Trigger_ID", "(Xx_)"
    InitializeMSVar ARES_NAME_LIBRARY_TYPE, "ARES_Library_Type_Name", "ARES"
    InitializeMSVar ARES_NAME_ITEM_TYPE, "ARES_Item_Type_Name", "ARESAutoLengthObject"
End Function

Private Function InitializeMSVar(ByRef msVar As ARES_MS_VAR, key As String, defaultValue As String)
    msVar.key = key
    msVar.Default = defaultValue
    msVar.Value = Config.GetVar(key)
    If msVar.Value = ARES_VAR.ARES_NAVD Then
        ResetMSVar msVar
    End If
End Function

Public Function ResetMSVar(ByRef msVar As ARES_MS_VAR)
    If Config.SetVar(msVar.key, msVar.Default) Then
        msVar.Value = Config.GetVar(msVar.key)
        ShowStatus msVar.key & " défini à " & msVar.Default & " par défaut"
    Else
        ShowStatus "Impossible de créer la variable " & msVar.key & " ou de la modifier."
    End If
End Function
