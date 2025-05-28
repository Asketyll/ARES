' Module: ARES_VAR
' Description: This module provides configuration variables, absolute values, and explanations of each variable and where they are used.

'######################################################################################################################
'                             CAN BE MODIFIED IN MS ENVIRONMENT VARIABLES DO NOT MODIFY HERE
'                                   USE CONFIG MODULE TO GET, SET AND DELETE A VALUE
'######################################################################################################################

' Used in Length module for default rounding
Public Const ARES_ROUNDS As String = "ARES_Round" 'Default Value: 2 (ARES_RND_DEFAULT)     Range 0 to 254 (Byte -1)    #255 it reserved for error (ARES_RND_ERROR_VALUE)

' Used in ElementChangeHandler ClassModule for automatically enable adding length to a text if the conditions are met.
Public Const ARES_AUTO_LENGTHS As String = "ARES_Auto_Lengths" 'Dafault Value: True (ARES_AUTO_LENGTHS_DEFAULT)     True or False (Boolean)

' Used in Auto_Lengths module For length-specific rounding in text.
Public Const ARES_LENGTH_ROUND As String = "ARES_Length_Round" 'Default Value: 1 (ARES_LENGTH_RND_DEFAULT) Range 0 to 254 (Byte -1) #255 it reserved for error (ARES_RND_ERROR_VALUE)

' Used in Auto_Lengths and StringsInEl module For Triggers in text.
Public Const ARES_LENGTH_TRIGGER As String = "ARES_Length_Triggers" 'Default Value: (Xx_m) (ARES_LENGTH_TRIGGER_DEFAULT) can a array use | (ARES_VAR_DELIMITER) like (Xx_m)|(Xx_cm)|(Xx_km)

' Used in AutoLengths and StringsInEl module for replace this triger with the length of element
Public Const ARES_LENGTH_TRIGGER_ID As String = "ARES_Length_Trigger_ID" 'Default Value: Xx_  (ARES_LENGTH_TRIGGER_ID_DEFAULT)

' Used in CustomPropertyHandler module for default name of Library Type object
Public Const ARES_NAME_LIBRARY_TYPE As String = "ARES_Library_Type"

' Used in CustomPropertyHandler module for default name of Item Type object
Public Const ARES_NAME_ITEM_TYPE As String = "ARES_Item_Type"


'######################################################################################################################
'                                   CAN'T BE EDITED IN MS VARIABLES ENVIRONMENT
'######################################################################################################################

' Used in Link module and ElementChangeHandler ClassModule for check if a graphics group exists
Public Const ARES_DEFAULT_GRAPHIC_GROUP_ID As Long = 0 ' Constant for no graphic group

' Used in Link and MicroStationDefinition module for check if a MsdElementType is unknow or raise a error
Public Const ARES_MSDETYPE_ERROR As Long = 44 ' If you use Type 44, you can replace with another MsdElementType not used

' Used in StringsInEl and AutoLengths module for separate a list in a environment variables
Public Const ARES_VAR_DELIMITER As String = "|" ' Delimiter used in :ARES_LENGTH_TRIGGER
