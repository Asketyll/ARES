' Module: ARES_VAR
' Description: This module provides configuration variables, absolute values, and explanations of each variable and where they are used.

'Used in Length module for default rounding
Public Const ROUNDS As String = "ARES_RND" 'Default Value: 2     Range 0 to 254 (Byte -1)    #255 it reserved for error

'Used in ElementChangeHandler ClassModule for automatically enable adding length to a text if the conditions are met.
Public Const AUTO_LENGTHS As String = "ARES_Auto_Lengths" 'Dafault Value: True      True or False (Boolean)

'Used in Auto_Lengths module For length-specific rounding in text.
Public Const LENGTH_ROUND As String = "ARES_Length_RND" 'Default Value: 1   Range 0 to 254 (Byte -1)    #255 it reserved for error

'Used in Auto_Lengths and StringsInEl module For Trigger in text.
Public Const LENGTH_TRIGGER As String = "ARES_Length_Trigger"

'Used in Link module and ElementChangeHandler ClassModule for check if a graphics group exists
Public Const DEFAULT_GRAPHIC_GROUP_ID As Long = 0 ' Constant for no graphic group

'Used in Link and MicroStationDefinition module for check if a MsdElementType is unknow or raise a error
Public Const MSDETYPE_ERROR As Long = 44 ' If you use Type 44, you can replace with another MsdElementType not used

'Used in AutoLengths module for replace this triger with the length of element
Public Const ARES_LENGTH_TRIGGER_ID As String = "Xx_"

'Used in CustomPropertyHandler module for default name of Library Type object
Public Const DEFAULT_NAME_LIBRARY_TYPE As String = "ARES_Library_Type"

'Used in CustomPropertyHandler module for default name of Item Type object
Public Const DEFAULT_NAME_ITEM_TYPE As String = "ARES_Item_Type"
