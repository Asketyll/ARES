' Module: ARESConstants
' Description: Contains all constants used in the ARES application
' License: This project is licensed under the AGPL-3.0.
' Dependencies: None
Option Explicit

'######################################################################################################################
'                    SYSTEM CONSTANTS - DO NOT MODIFY
'######################################################################################################################

' === GRAPHIC GROUP CONSTANTS ===
' Used in Link module and ElementChangeHandler ClassModule for check if a graphics group exists
Public Const ARES_DEFAULT_GRAPHIC_GROUP_ID As Long = 0

' === ELEMENT TYPE CONSTANTS ===
' Used in Link and MicroStationDefinition module for check if a MsdElementType is unknown or raise an error
' Note: If Type 44 is used elsewhere, replace with another unused MsdElementType
Public Const ARES_MSDETYPE_ERROR As Long = 44

' === STRING DELIMITER CONSTANTS ===
' Used in StringsInEl and AutoLengths module for separating lists in environment variables
Public Const ARES_VAR_DELIMITER As String = "|"

' === CONFIGURATION CONSTANTS ===
' Used in Config and ARESConfig modules
' Constant for undefined MS configuration variables (Not a Variable Defined)
Public Const ARES_NAVD As String = "NaVD"

' === ROUNDING ERROR CONSTANTS ===
' Reserved sentinel for all rounding config variables (ARES_ROUNDS, ARES_LENGTH_ROUND, ARES_ZONE_EXPORT_ROUND)
' Used in Length module and ExportLengthInRegion module
Public Const ARES_RND_ERROR_VALUE As Byte = 255

' === CELL CONSTANTS ===
' Used in CellRedraw module for error handling
Public Const ARES_CELL_INDEX_ERROR_VALUE As Integer = -1

' === CUSTOMPROPERTY LIB NAME CONSTANTS ===
' Used in CustomPropertyHandler module
Public Const ARES_NAME_LIBRARY_TYPE As String = "ARES"

' === CUSTOMPROPERTY ITEM NAME CONSTANTS ===
' Used in CustomPropertyHandler module
Public Const ARES_NAME_ITEM_TYPE As String = "ARESAutoLengthObject"

' === FILE DIALOG FILTER CONSTANTS ===
' Used in FileDialogs module — pipe-delimited Windows Forms filter strings
Public Const DIALOG_FILTER_CFG  As String = "ARES Config (*.cfg)|*.cfg|All Files (*.*)|*.*"
Public Const DIALOG_FILTER_XLSX As String = "Excel Workbook (*.xlsx)|*.xlsx|All Files (*.*)|*.*"

' === VERSION CONSTANTS ===
' Config schema version — written to exported .cfg files and checked on import
Public Const ARES_CONFIG_VERSION As String = "1.0.1"