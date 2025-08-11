' Module: ARESConstants
' Description: Contains all constants used in the ARES application
Option Explicit

'######################################################################################################################
'                       CAN'T BE EDITED IN MS VARIABLES ENVIRONMENT, YOU CAN MODIFY HERE
'######################################################################################################################
' Used in Link module and ElementChangeHandler ClassModule for check if a graphics group exists
Public Const ARES_DEFAULT_GRAPHIC_GROUP_ID As Long = 0 ' Constant for no graphic group
' Used in Link and MicroStationDefinition module for check if a MsdElementType is unknow or raise a error
Public Const ARES_MSDETYPE_ERROR As Long = 44 ' If you use Type 44, you can replace with another MsdElementType not used
' Used in StringsInEl and AutoLengths module for separate a list in a environment variables
Public Const ARES_VAR_DELIMITER As String = "|" ' Delimiter used in :ARES_LENGTH_TRIGGER
' Used in Config and ARESConfig
Public Const ARES_NAVD As String = "NaVD" ' Constant for undefined MS configuration variables :Not a Variable Defined
' Used in Length
Public Const ARES_RND_ERROR_VALUE As Byte = 255 ' Constant for error in ARES_ROUNDS and ARES_LENGTH_ROUND
' Used in CellRedreaw
Public Const ARES_CELL_INDEX_ERROR_VALUE As Integer = -1 ' Constant for error in CellRedreaw
