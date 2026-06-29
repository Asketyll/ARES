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

' === CUSTOM PROPERTY (ITEM TYPE) NAMING CONSTANTS ===
' Used in CustomPropertyHandler module.
' ARES stores ONE ItemType per custom property inside a single "ARES" ItemTypeLibrary, so each
' property can be attached independently to different elements. Each ItemType carries exactly one
' property, hence ItemType name == property name (that is what the MicroStation Properties dialog shows).
Public Const ARES_NAME_LIBRARY_TYPE As String = "ARES"          ' ItemTypeLibrary name (namespace)
Public Const ARES_ITEM_COMMUNE As String = "Commune"            ' ItemType carrying the editable text property
Public Const ARES_PROP_COMMUNE As String = "Commune"            ' free-text property (user typed)
Public Const ARES_ITEM_COUPE_TYPE As String = "Coupe Type"      ' ItemType carrying the value-list property
Public Const ARES_PROP_COUPE_TYPE As String = "Coupe Type"      ' value picked from ARES_Coupe_Type_List

' === FILE DIALOG FILTER CONSTANTS ===
' Used in FileDialogs module — pipe-delimited Windows Forms filter strings
Public Const DIALOG_FILTER_CFG  As String = "ARES Config (*.cfg)|*.cfg|All Files (*.*)|*.*"
Public Const DIALOG_FILTER_XLSX As String = "Excel Workbook (*.xlsx)|*.xlsx|All Files (*.*)|*.*"

' === VERSION CONSTANTS ===
' Config schema version — written to exported .cfg files and checked on import
Public Const ARES_CONFIG_VERSION As String = "1.0.1"

' === REGION SPLIT GEOMETRY CONSTANTS ===
' Used in RegionSplit module. Structural multipliers (dimensionless), NOT tolerance literals —
' the tolerances themselves are config vars (ARES_RegionSplit_Collinear_Tol / _Stroke_Tol).
' Knife half-width = collinear tolerance * this factor (sub-visible yet robust for GetRegionDifference).
Public Const ARES_KNIFE_HALFWIDTH_FACTOR As Double = 10#
' Knife half-width / over-extension floor as a fraction of the region bbox diagonal, so the slot
' stays above GetRegionDifference's extent-scaled cleanup tolerance at every scale (large regions
' otherwise collapse to a single region — "fewer than two regions (1)").
Public Const ARES_KNIFE_HALFWIDTH_REL_FACTOR As Double = 0.000005
' Knife over-extension past each chord end, as a fraction of the chord length (fully severs the region).
Public Const ARES_KNIFE_OVEREXTEND_FACTOR As Double = 0.01
' Minimum knife over-extension as a multiple of the stroke tolerance, so the knife clears the real
' curved boundary on an arc side (the stroked chord sits up to one stroke tolerance inside it).
Public Const ARES_KNIFE_ARC_OVEREXTEND_FACTOR As Double = 4#
' Lower / upper bound on the number of chords used to stroke an arc side into a polyline.
Public Const ARES_ARC_MIN_CHORDS As Long = 4
Public Const ARES_ARC_MAX_CHORDS As Long = 720

' === ZONING GEOMETRY CONSTANTS ===
' Used in Zoning module. A buffer piece endpoint within (buffer distance * this factor) of the
' chain's global Start/End point is treated as a free end (flat cap); any other endpoint is an
' interior junction (rounded cap). Coincident endpoints carry identical stored coordinates, so the
' real match distance is ~0; the factor only absorbs floating-point noise.
Public Const ARES_CAP_MATCH_FRAC As Double = 0.001