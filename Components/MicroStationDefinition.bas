' Module: MicroStationDefinition
' Description: This module provides functions to manipulate MsdElementType of MicroStation.
' It includes functions to convert strings to MsdElementType and validate MsdElementType values.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARES_VAR

Option Explicit

' Public function to convert a string to MsdElementType
Public Function StringToMsdElementType(ByVal TypeName As String, Optional CaseSensitive As Boolean = False) As MsdElementType
    On Error GoTo ErrorHandler
    
    Dim CompareMethod As VbCompareMethod
    Dim elementTypes As Object
    Dim ElementType As Variant
    Dim CompareResult As Integer

    ' Determine the comparison method based on CaseSensitive
    CompareMethod = IIf(CaseSensitive, vbBinaryCompare, vbTextCompare)

    ' Initialize a dictionary of MsdElementType values
    Set elementTypes = CreateObject("Scripting.Dictionary")
    InitializeElementTypes elementTypes

    ' Loop through the dictionary and compare typeName with each element type name
    For Each ElementType In elementTypes.Keys
        CompareResult = StrComp(TypeName, ElementType, CompareMethod)
        If CompareResult = 0 Then
            StringToMsdElementType = elementTypes(ElementType)
            Exit Function
        End If
    Next ElementType

ErrorHandler:
    ' Return error value in case of an error
    StringToMsdElementType = ARES_VAR.ARES_MSDETYPE_ERROR
	ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "MicroStationDefinition.StringToMsdElementType"
End Function

' Public function to check if a value is a valid MsdElementType
Public Function IsValidElementType(ByVal intValue As Integer) As Boolean
    On Error GoTo ErrorHandler

    ' Check if the value is within the range of MsdElementType enum
    If IsWithinValidRange(intValue) Then
        IsValidElementType = True
        Exit Function
    End If

ErrorHandler:
    ' Return False in case of an error
    IsValidElementType = False
	ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "MicroStationDefinition.IsValidElementType"
End Function

' Private Function to initialize the element types dictionary
Private Function InitializeElementTypes(ByRef elementTypes As Object)
    On Error GoTo ErrorHandler

    ' Add MsdElementType values to the dictionary
    elementTypes.Add "CellLibraryHeader", msdElementTypeCellLibraryHeader
    elementTypes.Add "CellHeader", msdElementTypeCellHeader
    elementTypes.Add "Line", msdElementTypeLine
    elementTypes.Add "LineString", msdElementTypeLineString
    elementTypes.Add "GroupData", msdElementTypeGroupData
    elementTypes.Add "Shape", msdElementTypeShape
    elementTypes.Add "TextNode", msdElementTypeTextNode
    elementTypes.Add "DigSetData", msdElementTypeDigSetData
    elementTypes.Add "DesignFileHeader", msdElementTypeDesignFileHeader
    elementTypes.Add "LevelSymbology", msdElementTypeLevelSymbology
    elementTypes.Add "Curve", msdElementTypeCurve
    elementTypes.Add "ComplexString", msdElementTypeComplexString
    elementTypes.Add "Conic", msdElementTypeConic
    elementTypes.Add "ComplexShape", msdElementTypeComplexShape
    elementTypes.Add "Ellipse", msdElementTypeEllipse
    elementTypes.Add "Arc", msdElementTypeArc
    elementTypes.Add "Text", msdElementTypeText
    elementTypes.Add "Surface", msdElementTypeSurface
    elementTypes.Add "Solid", msdElementTypeSolid
    elementTypes.Add "BsplinePole", msdElementTypeBsplinePole
    elementTypes.Add "PointString", msdElementTypePointString
    elementTypes.Add "Cone", msdElementTypeCone
    elementTypes.Add "BsplineSurface", msdElementTypeBsplineSurface
    elementTypes.Add "BsplineBoundary", msdElementTypeBsplineBoundary
    elementTypes.Add "BsplineKnot", msdElementTypeBsplineKnot
    elementTypes.Add "BsplineCurve", msdElementTypeBsplineCurve
    elementTypes.Add "BsplineWeight", msdElementTypeBsplineWeight
    elementTypes.Add "Dimension", msdElementTypeDimension
    elementTypes.Add "SharedCellDefinition", msdElementTypeSharedCellDefinition
    elementTypes.Add "SharedCell", msdElementTypeSharedCell
    elementTypes.Add "MultiLine", msdElementTypeMultiLine
    elementTypes.Add "Tag", msdElementTypeTag
    elementTypes.Add "DgnStoreComponent", msdElementTypeDgnStoreComponent
    elementTypes.Add "DgnStoreHeader", msdElementTypeDgnStoreHeader
    elementTypes.Add "44", msdElementType44
    elementTypes.Add "MicroStation", msdElementTypeMicroStation
    elementTypes.Add "RasterHeader", msdElementTypeRasterHeader
    elementTypes.Add "RasterComponent", msdElementTypeRasterComponent
    elementTypes.Add "RasterReference", msdElementTypeRasterReference
    elementTypes.Add "RasterReferenceComponent", msdElementTypeRasterReferenceComponent
    elementTypes.Add "RasterFrame", msdElementTypeRasterFrame
    elementTypes.Add "TableEntry", msdElementTypeTableEntry
    elementTypes.Add "Table", msdElementTypeTable
    elementTypes.Add "ViewGroup", msdElementTypeViewGroup
    elementTypes.Add "View", msdElementTypeView
    elementTypes.Add "LevelMask", msdElementTypeLevelMask
    elementTypes.Add "ReferenceAttachment", msdElementTypeReferenceAttachment
    elementTypes.Add "MatrixHeader", msdElementTypeMatrixHeader
    elementTypes.Add "MatrixIntegerData", msdElementTypeMatrixIntegerData
    elementTypes.Add "MatrixDoubleData", msdElementTypeMatrixDoubleData
    elementTypes.Add "MeshHeader", msdElementTypeMeshHeader
    elementTypes.Add "ReferenceOverride", msdElementTypeReferenceOverride
    elementTypes.Add "NamedGroupHeader", msdElementTypeNamedGroupHeader
    elementTypes.Add "NamedGroupComponent", msdElementTypeNamedGroupComponent

    Exit Function

ErrorHandler:
	ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "MicroStationDefinition.InitializeElementTypes"
End Function

' Private function to check if the value is within the valid range of MsdElementType
Private Function IsWithinValidRange(ByVal intValue As Integer) As Boolean
    On Error GoTo ErrorHandler

    ' Check if the value is within the valid range of MsdElementType enum
    IsWithinValidRange = (intValue >= msdElementTypeCellLibraryHeader And intValue <= msdElementTypeSolid) _
        Or (intValue >= msdElementTypeBsplinePole And intValue <= msdElementTypeBsplineWeight) _
        Or (intValue >= msdElementTypeDimension And intValue <= msdElementTypeDgnStoreHeader) _
        Or intValue = msdElementType44 _
        Or intValue = msdElementTypeMicroStation _
        Or (intValue >= msdElementTypeRasterHeader And intValue <= msdElementTypeRasterComponent) _
        Or (intValue >= msdElementTypeRasterReference And intValue <= msdElementTypeRasterReference) _
        Or (intValue >= msdElementTypeRasterFrame And intValue <= msdElementTypeMatrixDoubleData) _
        Or intValue = msdElementTypeMeshHeader _
        Or intValue = msdElementTypeReferenceOverride _
        Or intValue = msdElementTypeNamedGroupHeader _
        Or intValue = msdElementTypeNamedGroupComponent

    Exit Function

ErrorHandler:
    ' Return False in case of an error
    IsWithinValidRange = False
	ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "MicroStationDefinition.IsWithinValidRange"
End Function
