' Module: CellRedreaw
' Description: This module provides functions to update and manage ATLAS cell labels in MicroStation.
' It includes functions to calculate the size and origin of text elements, find connector indices between elements,
' move vertices and elements, and rotate elements within a cell. The module ensures that operations are performed
' with silent error handling to avoid interrupting the workflow.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, ARESConstants, ErrorHandlerClass

Option Explicit

Public Delta(4) As Integer

'work with ETI01E, ETI01F, ETI01K, ETI01M, ETI01N, ETI01O and ETI076 cell of ATLAS lib

' Constants
Private Const DEFAULT_TOLERANCE_RATIO As Double = 50
Private PI_OVER_2 As Double
Private PI_OVER_7 As Double

Public Function ATLASCellLabelUpdate(El As element) As Boolean
    On Error GoTo ErrorHandler
    
    PI_OVER_2 = Application.Pi / 2
    PI_OVER_7 = Application.Pi / 7
    
    Dim LinesIndex() As Integer
    Dim RectanglesIndex() As Integer
    Dim TextSize() As Double
    Dim TextOrigine As Point3d
    Dim ShapesIndexByArea() As Integer
    Dim ConnectShapeAndLine() As Integer
    Dim ConnectBigShapeAndMediumShape() As Integer
    Dim ConnectMediumShapeAndTinyShape() As Integer
    Dim AngleRad As Double
    Dim Corners(4) As Point3d
    Dim PostAngle As Double
    Dim ShapeVertices() As Point3d
    Dim Delta1 As Point3d
    Dim Delta2 As Point3d
    
    If Not CheckInitialConditions(El) Then Exit Function
    
    ' Refresh the element reference
    RefreshoCell El
    
    ' Get the indices of line elements within the cell
    LinesIndex = GetLineIndicesInCell(El)
    If UBound(LinesIndex) < 0 Then Exit Function
    
    ' Get the indices of rectangle elements within the cell
    RectanglesIndex = GetRectangleIndicesInCell(El)
    If UBound(RectanglesIndex) < 0 Then Exit Function
    
    ' Get text size and origin
    TextSize = GetTextSize(El)
    If TextSize(0) <= 0 Or TextSize(1) <= 0 Then Exit Function
    
    TextOrigine = GetTextOrigin(El)
    If Point3dEqual(TextOrigine, Point3dFromXY(0, 0)) Then Exit Function
    
    ' Get the indices of quadrilateral shapes within the cell, sorted by area
    ShapesIndexByArea = GetQuadrilateralIndicesByArea(El)
    If UBound(ShapesIndexByArea) < 2 Then Exit Function
    
    ' Get connectors
    ConnectShapeAndLine() = GetConnector(GetElementAtIndex(El, ShapesIndexByArea(0)), GetElementAtIndex(El, LinesIndex(0)))
    If ConnectShapeAndLine(0) < 0 Or ConnectShapeAndLine(1) < 0 Then Exit Function
    
    ConnectBigShapeAndMediumShape() = GetConnector(GetElementAtIndex(El, ShapesIndexByArea(0)), GetElementAtIndex(El, ShapesIndexByArea(1)))
    If ConnectBigShapeAndMediumShape(0) < 0 Or ConnectBigShapeAndMediumShape(1) < 0 Then Exit Function
    
    ConnectMediumShapeAndTinyShape() = GetConnector(GetElementAtIndex(El, ShapesIndexByArea(1)), GetElementAtIndex(El, ShapesIndexByArea(2)))
    If ConnectMediumShapeAndTinyShape(0) < 0 Or ConnectMediumShapeAndTinyShape(1) < 0 Then Exit Function
    
    ' Determine the rotation angle of the cell
    Matrix3dIsXYRotation El.AsCellElement.Rotation, AngleRad
    
    ' Adjust the text size
    TextSize(0) = TextSize(0) + (TextSize(1) * 1.5)
    TextSize(1) = TextSize(1) + (TextSize(1) * 1.5)
    
    ' Calculate the corners of the text area
    If Not CalculateCorners(TextOrigine, TextSize, AngleRad, Corners) Then Exit Function
    
    ' Determine the post angle based on cell name
    Select Case True
        Case El.AsCellElement.Name = "ETI076"
            PostAngle = PI_OVER_2
        Case Else
            PostAngle = PI_OVER_7
    End Select
    
    ' Define the order of the corners, depending of cell
    ShapeVertices = GetElementAtIndex(El, ShapesIndexByArea(0)).AsShapeElement.GetVertices
    Delta(0) = FindClosestVertex(ShapeVertices(0), Corners)
    Delta(1) = FindClosestVertex(ShapeVertices(1), Corners)
    Delta(2) = FindClosestVertex(ShapeVertices(2), Corners)
    Delta(3) = FindClosestVertex(ShapeVertices(3), Corners)
    Delta(4) = FindClosestVertex(ShapeVertices(4), Corners)
    
    ' Calculate coordinate deltas for adjusting the positions of shapes
    Delta1 = CalculateCoordinateDelta(GetElementAtIndex(El, ShapesIndexByArea(0)), Delta(3), GetElementAtIndex(El, ShapesIndexByArea(1)), 0)
    If Point3dEqual(Delta1, Point3dFromXY(0, 0)) Then Exit Function
    
    Delta2 = CalculateCoordinateDelta(GetElementAtIndex(El, ShapesIndexByArea(0)), Delta(3), GetElementAtIndex(El, ShapesIndexByArea(1)), 3)
    If Point3dEqual(Delta2, Point3dFromXY(0, 0)) Then Exit Function
    
    ' Update the shape in the cell with new vertices and rotation
    If Not UpdateShapes(El, ShapesIndexByArea, Corners, TextOrigine, LinesIndex, RectanglesIndex, Delta1, Delta2, _
                        ConnectShapeAndLine, ConnectBigShapeAndMediumShape, ConnectMediumShapeAndTinyShape, PostAngle) Then Exit Function
    
    ATLASCellLabelUpdate = True
    Exit Function
    
ErrorHandler:
    ErrorHandler.LogError Err.Description, "CellRedraw.ATLASCellLabelUpdate"
    ATLASCellLabelUpdate = False
End Function

Private Function UpdateShapes(El As element, ShapesIndexByArea() As Integer, Corners() As Point3d, _
                              TextOrigine As Point3d, LinesIndex() As Integer, RectanglesIndex() As Integer, _
                              Delta1 As Point3d, Delta2 As Point3d, ConnectShapeAndLine() As Integer, _
                              ConnectBigShapeAndMediumShape() As Integer, ConnectMediumShapeAndTinyShape() As Integer, _
                              PostAngle As Double) As Boolean
    On Error GoTo ErrorHandler
    
    Dim LineOrigin As Point3d
    Dim BigShapeEl As element
    Dim MediumShapeEl As element
    Dim BigShapeVertex As Point3d
    Dim NewCoordinate1 As Point3d
    Dim NewCoordinate2 As Point3d
    
    ' Update the main shape
    If Not UpdateShapeInCell(El, ShapesIndexByArea(0), Corners, TextOrigine) Then Exit Function
    
    ' Move a vertex of the line to match a vertex of the shape
    If Not MoveVertexToVertexInCell(El, ShapesIndexByArea(0), LinesIndex(0), ConnectShapeAndLine) Then Exit Function
    
    ' Move the medium shape to match a vertex of the big shape
    If Not MoveElementToVertexInCell(El, ShapesIndexByArea(0), ShapesIndexByArea(1), ConnectBigShapeAndMediumShape) Then Exit Function
    
    ' Move the tiny shape to match a vertex of the medium shape
    If Not MoveElementToVertexInCell(El, ShapesIndexByArea(1), ShapesIndexByArea(2), ConnectMediumShapeAndTinyShape) Then Exit Function
    
    ' Rotate the rectangle element like the line element
    LineOrigin = GetCoordinate(GetElementAtIndex(El, LinesIndex(0)), 0)
    If Not RotateElementLikeElementInCell(El, RectanglesIndex(0), LinesIndex(0), LineOrigin, PostAngle) Then Exit Function
    
    ' Move vertices of the medium shape to new coordinates
    Set BigShapeEl = GetElementAtIndex(El, ShapesIndexByArea(0))
    Set MediumShapeEl = GetElementAtIndex(El, ShapesIndexByArea(1))
    
    BigShapeVertex = GetCoordinate(BigShapeEl, Delta(3))
    
    NewCoordinate1 = Point3dSubtract(BigShapeVertex, Delta1)
    NewCoordinate2 = Point3dSubtract(BigShapeVertex, Delta2)

    If Not MoveVertexToCoordinateInCell(El, ShapesIndexByArea(1), 0, NewCoordinate1) Then Exit Function
    If Not MoveVertexToCoordinateInCell(El, ShapesIndexByArea(1), 4, NewCoordinate1) Then Exit Function
    If Not MoveVertexToCoordinateInCell(El, ShapesIndexByArea(1), 3, NewCoordinate2) Then Exit Function

    UpdateShapes = True
    Exit Function

ErrorHandler:
    ErrorHandler.LogError Err.Description, "CellRedraw.UpdateShapes"
    UpdateShapes = False
End Function

Private Function CheckInitialConditions(El As element) As Boolean
    On Error GoTo ErrorHandler
    
    Dim CellsName() As String
    Dim i As Long
    
    ' Check if the configuration allows updating ATLAS cell labels
    If ARESConfig.ARES_UPDATE_ATLASCELLLABEL.Value Then
    Else
        Exit Function
    End If
    ' Check if the element is a cell element
    If Not El.IsCellElement Then Exit Function
    
    ' Split the configuration string to get the list of cell names that should be updated
    CellsName = Split(ARESConfig.ARES_CELL_LIKE_LABEL.Value, ARESConstants.ARES_VAR_DELIMITER)
    
    ' Loop through each cell name to check if the current element matches
    For i = LBound(CellsName) To UBound(CellsName)
        If El.AsCellElement.Name = CellsName(i) Then
            CheckInitialConditions = True
            Exit Function
        End If
    Next i

ErrorHandler:
    ErrorHandler.LogError Err.Description, "CellRedraw.CheckInitialConditions"
    CheckInitialConditions = False
End Function

' Function to get the indices of LineElements within a CellElement
' Returns an array of indices, an empty array if no LineElement is found, or an array with ARESConstants.ARES_CELL_INDEX_ERROR_VALUE if an error occurs (-1)
Private Function GetLineIndicesInCell(CellEl As CellElement) As Integer()
    On Error GoTo ErrorHandler
    
    Dim Indices() As Integer
    Dim indexCount As Integer
    
    ' Initialize Indices as an empty integer array
    ReDim Indices(-1 To -1)
    
    indexCount = -1 ' Start at -1 to account for zero-based index
    
    ' Reset the enumeration to start from the first element
    CellEl.ResetElementEnumeration
    
    ' Loop through each element in the CellElement
    Do While CellEl.MoveToNextElement
        indexCount = indexCount + 1
        If CellEl.CopyCurrentElement.IsLineElement Then
            ' Resize the array and store the index
            If UBound(Indices) = -1 Then
                ReDim Indices(0)
            Else
                ReDim Preserve Indices(UBound(Indices) + 1)
            End If
            Indices(UBound(Indices)) = indexCount
        End If
    Loop
    
    ' Return the indices array
    GetLineIndicesInCell = Indices
    Exit Function
    
ErrorHandler:
    ErrorHandler.LogError Err.Description, "CellRedraw.GetLineIndicesInCell"
    Dim ErrorIndices(-1 To 0) As Integer
    ErrorIndices(0) = ARESConstants.ARES_CELL_INDEX_ERROR_VALUE
    GetLineIndicesInCell = ErrorIndices
End Function

' Function to get the indices of rectangular ShapeElements within a CellElement
' Returns an array of indices, an empty array if no such ShapeElement is found, or an array with ARESConstants.ARES_CELL_INDEX_ERROR_VALUE if an error occurs (-1)
Private Function GetRectangleIndicesInCell(CellEl As CellElement) As Integer()
    On Error GoTo ErrorHandler
    
    Dim Indices() As Integer
    Dim indexCount As Integer
    
    ' Initialize Indices as an empty integer array
    ReDim Indices(-1 To -1)
    
    indexCount = -1 ' Start at -1 to account for zero-based index
    
    ' Reset the enumeration to start from the first element
    CellEl.ResetElementEnumeration
    
    ' Loop through each element in the CellElement
    Do While CellEl.MoveToNextElement
        indexCount = indexCount + 1
        If CellEl.CopyCurrentElement.IsShapeElement Then
            If CellEl.CopyCurrentElement.AsShapeElement.VerticesCount - 1 = 3 Then
                ' Resize the array and store the index
                If UBound(Indices) = -1 Then
                    ReDim Indices(0)
                Else
                    ReDim Preserve Indices(UBound(Indices) + 1)
                End If
                Indices(UBound(Indices)) = indexCount
            End If
        End If
    Loop
    
    ' Return the indices array
    GetRectangleIndicesInCell = Indices
    Exit Function
    
ErrorHandler:
    ErrorHandler.LogError Err.Description, "CellRedraw.GetRectangleIndicesInCell"
    Dim ErrorIndices(-1 To 0) As Integer
    ErrorIndices(0) = ARESConstants.ARES_CELL_INDEX_ERROR_VALUE
    GetRectangleIndicesInCell = ErrorIndices
End Function

' Function to get the size of the text within a CellElement
' Returns an array of doubles representing the text size, or an array with ARESConstants.ARES_CELL_INDEX_ERROR_VALUE if an error occurs (-1)
Private Function GetTextSize(CellEl As CellElement) As Double()
    On Error GoTo ErrorHandler
    
    Dim TextSize(1) As Double
    Dim LocalTextSize(1) As Double
    Dim ee As ElementEnumerator
    Dim subel As TextElement

    ' Initialize TextSize to zero
    TextSize(0) = 0
    TextSize(1) = 0
    
    ' Reset the enumeration to start from the first element
    CellEl.ResetElementEnumeration
    
    ' Loop through each element in the CellElement
    Do While CellEl.MoveToNextElement
        If CellEl.CopyCurrentElement.IsTextElement Or CellEl.CopyCurrentElement.IsTextNodeElement Then
            Select Case True
                Case CellEl.CopyCurrentElement.IsTextElement
                    ' Get the total text size for TextElement
                    CellEl.CopyCurrentElement.AsTextElement.GetTotalTextSize TextSize(0), TextSize(1)
                    GetTextSize = TextSize
                    Exit Function
                Case CellEl.CopyCurrentElement.IsTextNodeElement
                    ' Get the total text size for TextNodeElement
                    Set ee = CellEl.CopyCurrentElement.AsTextNodeElement.GetSubElements
                    While ee.MoveNext
                        Set subel = ee.Current
                        subel.GetTotalTextSize LocalTextSize(0), LocalTextSize(1)
                        TextSize(1) = TextSize(1) + LocalTextSize(1)
                        If LocalTextSize(0) > TextSize(0) Then
                            TextSize(0) = LocalTextSize(0)
                        End If
                    Wend
                    GetTextSize = TextSize
                    Exit Function
            End Select
            Exit Do
        End If
    Loop

    Exit Function

ErrorHandler:
    ' Log the error and return an array with error values
    ErrorHandler.LogError Err.Description, "CellRedraw.GetTextSize"
    GetTextSize = Array(CDbl(ARESConstants.ARES_CELL_INDEX_ERROR_VALUE), CDbl(ARESConstants.ARES_CELL_INDEX_ERROR_VALUE))
End Function

' Function to get the origin of the text within a CellElement
' Returns a Point3d representing the text origin, or a Point3d with zero values if an error occurs
Private Function GetTextOrigin(CellEl As CellElement) As Point3d
    On Error GoTo ErrorHandler

    Dim origin As Point3d

    ' Reset the enumeration to start from the first element
    CellEl.ResetElementEnumeration

    ' Loop through each element in the CellElement
    Do While CellEl.MoveToNextElement
        If CellEl.CopyCurrentElement.IsTextElement Or CellEl.CopyCurrentElement.IsTextNodeElement Then
            Select Case True
                Case CellEl.CopyCurrentElement.IsTextElement
                    ' Get the origin for TextElement
                    origin = CellEl.CopyCurrentElement.AsTextElement.origin
                    GetTextOrigin = origin
                    Exit Function
                Case CellEl.CopyCurrentElement.IsTextNodeElement
                    ' Get the origin for TextNodeElement
                    origin = CellEl.CopyCurrentElement.AsTextNodeElement.origin
                    GetTextOrigin = origin
                    Exit Function
            End Select
        End If
    Loop

    Exit Function

ErrorHandler:
    ' Log the error and return a Point3d with zero values
    ErrorHandler.LogError Err.Description, "CellRedraw.GetTextOrigin"
    GetTextOrigin = Point3dZero
End Function

' Function to get the coordinate of a vertex in an Element
' Returns a Point3d representing the coordinate, or a Point3d with zero values if an error occurs
Private Function GetCoordinate(El As element, index As Integer) As Point3d
    On Error GoTo ErrorHandler
    
    Dim Vertices() As Point3d

    ' Check if the index is valid
    If index < 0 Or index = ARESConstants.ARES_CELL_INDEX_ERROR_VALUE Then
        GetCoordinate = Point3dZero
        Exit Function
    End If
    
    ' Get the vertices based on the element type
    Select Case True
        Case El.IsLineElement
            Vertices = El.AsLineElement.GetVertices
        Case El.IsShapeElement
            Vertices = El.AsShapeElement.GetVertices
        Case Else
            ' Unsupported element type
            GetCoordinate = Point3dZero
            Exit Function
    End Select
    
    ' Check if the index is within the valid range
    If index > UBound(Vertices) Then
        GetCoordinate = Point3dZero
        Exit Function
    End If
    
    ' Return the coordinate at the specified index
    GetCoordinate = Vertices(index)
    Exit Function

ErrorHandler:
    ' Log the error and return a Point3d with error values
    ErrorHandler.LogError Err.Description, "ElementUtils.GetCoordinate"
    GetCoordinate = Point3dZero
End Function

' Function to get the indices of quadrilateral ShapeElements within a CellElement, sorted by area in descending order.
' Returns an empty array if no quadrilateral ShapeElements are found or an array with ARESConstants.ARES_CELL_INDEX_ERROR_VALUE if an error occurs.
' Can't return more than 255 indices; use ElementEnumerator if you need more.
Private Function GetQuadrilateralIndicesByArea(CellEl As CellElement) As Integer()
    On Error GoTo ErrorHandler

    Dim Areas() As Double
    Dim Indices() As Integer
    Dim ElementCount As Integer

    ' First pass: Count the number of quadrilateral ShapeElements
    ElementCount = CountQuadrilaterals(CellEl)
    If ElementCount = ARESConstants.ARES_CELL_INDEX_ERROR_VALUE Then
        GetQuadrilateralIndicesByArea = Array(ARESConstants.ARES_CELL_INDEX_ERROR_VALUE)
        Exit Function
    End If

    ' If no quadrilateral ShapeElements found, return an empty array
    If ElementCount = 0 Then
        GetQuadrilateralIndicesByArea = Array()
        Exit Function
    End If

    ' Initialize arrays to store areas and indices of quadrilateral ShapeElements
    ReDim Areas(ElementCount - 1)
    ReDim Indices(ElementCount - 1)

    ' Second pass: Collect areas and indices of quadrilateral ShapeElements
    ElementCount = CollectQuadrilateralData(CellEl, Areas, Indices, ElementCount)
    If ElementCount = ARESConstants.ARES_CELL_INDEX_ERROR_VALUE Then
        GetQuadrilateralIndicesByArea = Array(ARESConstants.ARES_CELL_INDEX_ERROR_VALUE)
        Exit Function
    End If

    ' Sort the areas and indices in descending order of area
    If Not SortAreasAndIndices(Areas, Indices, ElementCount) Then
        GetQuadrilateralIndicesByArea = Array(ARESConstants.ARES_CELL_INDEX_ERROR_VALUE)
        Exit Function
    End If

    ' Return the sorted indices of quadrilateral ShapeElements
    GetQuadrilateralIndicesByArea = Indices
    Exit Function

ErrorHandler:
    ' Log the error and return an array with error values
    ErrorHandler.LogError Err.Description, "CellElementUtils.GetQuadrilateralIndicesByArea"
    GetQuadrilateralIndicesByArea = Array(ARESConstants.ARES_CELL_INDEX_ERROR_VALUE)
End Function

' Function to count the number of quadrilateral ShapeElements in a CellElement
' Can't return more than 255 indices; use ElementEnumerator if you need more.
Private Function CountQuadrilaterals(CellEl As CellElement) As Integer
    On Error GoTo ErrorHandler

    Dim Vertices() As Point3d
    Dim ElementCount As Integer
    Dim CurrentIndex As Integer

    ElementCount = 0
    CurrentIndex = 0

    CellEl.ResetElementEnumeration
    Do While CellEl.MoveToNextElement
        If CellEl.CopyCurrentElement.IsShapeElement Then
            Vertices = CellEl.CopyCurrentElement.AsShapeElement.GetVertices
            If UBound(Vertices) > 3 Then
                ElementCount = ElementCount + 1
            End If
        End If
        CurrentIndex = CurrentIndex + 1
        If CurrentIndex > 255 Then Exit Do
    Loop

    CountQuadrilaterals = ElementCount
    Exit Function

ErrorHandler:
    ' Log the error and return an error value
    ErrorHandler.LogError Err.Description, "CellElementUtils.CountQuadrilaterals"
    CountQuadrilaterals = ARESConstants.ARES_CELL_INDEX_ERROR_VALUE
End Function

' Function to collect areas and indices of quadrilateral ShapeElements in a CellElement
' Can't return more than 255 indices; use ElementEnumerator if you need more.
Private Function CollectQuadrilateralData(CellEl As CellElement, Areas() As Double, Indices() As Integer, ElementCount As Integer) As Integer
    On Error GoTo ErrorHandler

    Dim Vertices() As Point3d
    Dim CurrentIndex As Integer
    Dim Area As Double
    Dim Count As Integer

    Count = 0
    CurrentIndex = 0

    CellEl.ResetElementEnumeration
    Do While CellEl.MoveToNextElement
        If CellEl.CopyCurrentElement.IsShapeElement Then
            Vertices = CellEl.CopyCurrentElement.AsShapeElement.GetVertices
            If UBound(Vertices) > 3 Then
                Area = CellEl.CopyCurrentElement.AsShapeElement.Area
                Areas(Count) = Area
                Indices(Count) = CurrentIndex
                Count = Count + 1
            End If
        End If
        CurrentIndex = CurrentIndex + 1
        If CurrentIndex > 255 Then Exit Do
    Loop

    CollectQuadrilateralData = Count
    Exit Function

ErrorHandler:
    ' Log the error and return an error value
    ErrorHandler.LogError Err.Description, "CellElementUtils.CollectQuadrilateralData"
    CollectQuadrilateralData = ARESConstants.ARES_CELL_INDEX_ERROR_VALUE
End Function

' Function to sort areas and indices in descending order of area
' Can't return more than 255 indices; use ElementEnumerator if you need more.
Private Function SortAreasAndIndices(Areas() As Double, Indices() As Integer, ElementCount As Integer) As Boolean
    On Error GoTo ErrorHandler

    Dim i As Integer, j As Integer
    Dim TempArea As Double
    Dim TempIndex As Integer

    For i = 0 To ElementCount - 2
        For j = i + 1 To ElementCount - 1
            If Areas(i) < Areas(j) Then
                ' Swap areas
                TempArea = Areas(i)
                Areas(i) = Areas(j)
                Areas(j) = TempArea

                ' Swap indices
                TempIndex = Indices(i)
                Indices(i) = Indices(j)
                Indices(j) = TempIndex
            End If
        Next j
    Next i

    SortAreasAndIndices = True
    Exit Function

ErrorHandler:
    ' Log the error and return False to indicate failure
    ErrorHandler.LogError Err.Description, "CellElementUtils.SortAreasAndIndices"
    SortAreasAndIndices = False
End Function

' Function to get the element at a specific index within a CellElement.
' If you have much sub element in your cell, please use ElementEnumerator
Private Function GetElementAtIndex(CellEl As CellElement, index As Integer) As element
    On Error GoTo ErrorHandler
    
    Dim elementsArray() As element
    
    ' Get the sub-elements of the CellElement as an array
    elementsArray = CellEl.GetSubElements.BuildArrayFromContents
    
    ' Check if the index is within the valid range
    If index >= 0 And index <= UBound(elementsArray) Then
        ' Return the element at the specified index
        Set GetElementAtIndex = elementsArray(index)
        Exit Function
    End If

    Set GetElementAtIndex = Nothing
    Exit Function

ErrorHandler:
    ErrorHandler.LogError Err.Description, "CellRedreaw.GetElementAtIndex"
    Set GetElementAtIndex = Nothing
End Function

Private Function RefreshoCell(oCell As CellElement)
    Set oCell = ActiveDesignFile.GetElementByID(oCell.id)
End Function

' Function to update a ShapeElement within a CellElement and ensure the CellElement is up-to-date
Private Function UpdateShapeInCell(CellEl As CellElement, ShapeIndex As Integer, NewVertices() As Point3d, oOrigin As Point3d) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ShapeEl As ShapeElement
    Dim i As Long
    
    ' Get the ShapeElement from the CellElement using the provided index
    Set ShapeEl = GetElementAtIndex(CellEl, ShapeIndex)
    If ShapeEl Is Nothing Then
        UpdateShapeInCell = False
        Exit Function
    End If
    
    ' Check if the number of new vertices matches the vertices count of the ShapeElement
    If UBound(NewVertices) - LBound(NewVertices) + 1 <> ShapeEl.VerticesCount Then
        UpdateShapeInCell = False
        Exit Function
    End If
    
    ' Modify the vertices of the ShapeElement
    For i = 0 To ShapeEl.VerticesCount - 1
        ShapeEl.ModifyVertex Delta(i), NewVertices(i)
    Next i
    ShapeEl.Rewrite
    
    ' Rewrite the ShapeEl to apply the changes
    ShapeEl.Rewrite
    
    ' Refresh object for refresh sub element
    RefreshoCell CellEl
    
    ' Return True to indicate success
    UpdateShapeInCell = True
    Exit Function

ErrorHandler:
    ErrorHandler.LogError Err.Description, "CellRedreaw.UpdateShapeInCell"
    UpdateShapeInCell = False
End Function

' Function to calculate the corners of a rectangle based on its origin and size.
Private Function CalculateCorners(oOrigine As Point3d, TextSize() As Double, AngleRad As Double, Corners() As Point3d) As Boolean
    On Error GoTo ErrorHandler

    ' Check if TextSize and Corners arrays have the correct dimensions
    If UBound(TextSize) < 1 Or UBound(Corners) < 4 Then
        CalculateCorners = False
        Exit Function
    End If
    
    Dim halfWidth As Double
    Dim halfHeight As Double
    
    ' Calculate half of the width and height
    halfWidth = TextSize(0) / 2
    halfHeight = TextSize(1) / 2
    
    ' Calculate the corners of the rectangle
    Corners(0).X = oOrigine.X - halfWidth
    Corners(0).Y = oOrigine.Y + halfHeight
    Corners(1).X = oOrigine.X + halfWidth
    Corners(1).Y = oOrigine.Y + halfHeight
    Corners(2).X = oOrigine.X + halfWidth
    Corners(2).Y = oOrigine.Y - halfHeight
    Corners(3).X = oOrigine.X - halfWidth
    Corners(3).Y = oOrigine.Y - halfHeight
    Corners(4).X = oOrigine.X - halfWidth
    Corners(4).Y = oOrigine.Y + halfHeight
    
    Dim oSquareElt As ShapeElement
    Set oSquareElt = CreateShapeElement1(Nothing, Corners, msdFillModeNotFilled)
    oSquareElt.RotateAboutZ oOrigine, AngleRad
    Dim Vertices() As Point3d
    Dim i As Integer
    Vertices = oSquareElt.GetVertices
    
    For i = 0 To UBound(Corners)
        Corners(i) = Vertices(i)
    Next i
    ' Return True to indicate success
    CalculateCorners = True

    Exit Function
    
ErrorHandler:
    ' Log the error and return False to indicate failure
    ErrorHandler.LogError Err.Description, "CellRedreaw.CalculateCorners"
    CalculateCorners = False
End Function

' Function to find the connector indices between a master ShapeElement and a child Element.
Private Function GetConnector(MasterEl As ShapeElement, ChidrenEl As element, Optional ToleranceDependingOnSizeOfMaster As Double) As Integer()
    On Error GoTo ErrorHandler
    
    Dim MasterVertices() As Point3d
    Dim ChidrenVertices() As Point3d
    Dim i As Byte
    Dim j As Byte
    Dim Connector(1) As Integer
    
    ' Initialize the connector array with default error values
    Connector(0) = ARESConstants.ARES_CELL_INDEX_ERROR_VALUE
    Connector(1) = ARESConstants.ARES_CELL_INDEX_ERROR_VALUE
    
    ' Get the vertices of the master ShapeElement
    MasterVertices = MasterEl.GetVertices
    
    ' Get the vertices of the child Element based on its type
    Select Case True
        Case ChidrenEl.IsShapeElement
            ChidrenVertices = ChidrenEl.AsShapeElement.GetVertices
        Case ChidrenEl.IsLineElement
            ChidrenVertices = ChidrenEl.AsLineElement.GetVertices
        Case Else
            ' Unsupported child element type
            GetConnector = Connector
            Exit Function
    End Select
    
    ' Calculate the tolerance if not provided
    If ToleranceDependingOnSizeOfMaster = 0 Then
        ToleranceDependingOnSizeOfMaster = Point3dDistance(MasterVertices(0), MasterVertices(3)) 'Never absolut Value if you want work with any drawing scale
        ToleranceDependingOnSizeOfMaster = ToleranceDependingOnSizeOfMaster / DEFAULT_TOLERANCE_RATIO 'Arbitrary value
    End If
    
    ' Loop through the vertices of the master and child elements to find a connector
    For i = 0 To UBound(MasterVertices)
        For j = 0 To UBound(ChidrenVertices)
            If Point3dEqualTolerance(MasterVertices(i), ChidrenVertices(j), ToleranceDependingOnSizeOfMaster) Then
                Connector(0) = i
                Connector(1) = j
                GetConnector = Connector
                Exit Function
            End If
        Next j
    Next i
    
    ' If no connector is found, return the default error values
    GetConnector = Connector
    Exit Function
    
ErrorHandler:
    ' Log the error and return an array with error values
    ErrorHandler.LogError Err.Description, "CellRedreaw.GetConnector"
    GetConnector = Connector
End Function

'Function to move a vertex of the child element to a vertex of the master element within a cell.
Private Function MoveVertexToVertexInCell(CellEl As CellElement, MasterIndex As Integer, ChildIndex As Integer, index() As Integer) As Boolean
    On Error GoTo ErrorHandler
    
    Dim MasterEl As element
    Dim ChidrenEl As element
    Dim MasterVertices() As Point3d
    Dim ChidrenVertices() As Point3d
    Dim NewPoint As Point3d
    
    ' Get the master and child elements from the CellElement using the provided indices
    Set MasterEl = GetElementAtIndex(CellEl, MasterIndex)
    If MasterEl Is Nothing Then Exit Function

    Set ChidrenEl = GetElementAtIndex(CellEl, ChildIndex)
    If ChidrenEl Is Nothing Then Exit Function
    
    ' Check if the master element is a ShapeElement
    If Not MasterEl.IsShapeElement Then Exit Function

    ' Get the vertices of the master ShapeElement
    MasterVertices = MasterEl.AsShapeElement.GetVertices
    
    ' Get the vertices of the child element based on its type
    Select Case True
        Case ChidrenEl.IsLineElement
            ChidrenVertices = ChidrenEl.AsLineElement.GetVertices
        Case ChidrenEl.IsShapeElement
            ChidrenVertices = ChidrenEl.AsShapeElement.GetVertices
        Case Else
            ' Unsupported child element type
            Exit Function
    End Select
    
    ' Check if the provided indices are valid
    If index(0) >= 0 And index(0) <= UBound(MasterVertices) And _
       index(1) >= 0 And index(1) <= UBound(ChidrenVertices) Then

        ' Get the new point from the master element's vertex
        NewPoint = MasterVertices(index(0))

        ' Modify the vertex of the child element
        Select Case True
            Case ChidrenEl.IsLineElement
                ChidrenEl.AsLineElement.ModifyVertex index(1), NewPoint
            Case ChidrenEl.IsShapeElement
                ChidrenEl.AsShapeElement.ModifyVertex index(1), NewPoint
        End Select

        ChidrenEl.Rewrite

        ' Refresh object for refresh sub element
        RefreshoCell CellEl
        
        ' The operation was successful
        MoveVertexToVertexInCell = True
        Exit Function
    End If

    Exit Function

ErrorHandler:
    ErrorHandler.LogError Err.Description, "CellRedreaw.MoveVertexToVertexInCell"
    MoveVertexToVertexInCell = False
End Function

'Function to move a element(child element) to a vertex of the master element within a cell.
Private Function MoveElementToVertexInCell(CellEl As CellElement, MasterIndex As Integer, ChildIndex As Integer, index() As Integer) As Boolean
    On Error GoTo ErrorHandler

    Dim MasterEl As element
    Dim ChidrenEl As element
    Dim MasterVertices() As Point3d
    Dim ChidrenVertices() As Point3d
    Dim offset As Point3d

    ' Get the master and child elements from the CellElement using the provided indices
    Set MasterEl = GetElementAtIndex(CellEl, MasterIndex)
    If MasterEl Is Nothing Then Exit Function

    Set ChidrenEl = GetElementAtIndex(CellEl, ChildIndex)
    If ChidrenEl Is Nothing Then Exit Function
    
    ' Get the vertices of the master ShapeElement
    MasterVertices = MasterEl.AsShapeElement.GetVertices
    
    ' Get the vertices of the child element based on its type
    Select Case True
        Case ChidrenEl.IsLineElement
            ChidrenVertices = ChidrenEl.AsLineElement.GetVertices
        Case ChidrenEl.IsShapeElement
            ChidrenVertices = ChidrenEl.AsShapeElement.GetVertices
        Case Else
            ' Unsupported child element type
            Exit Function
    End Select

    ' Check if the provided indices are valid
    If index(0) >= 0 And index(0) <= UBound(MasterVertices) And _
       index(1) >= 0 And index(1) <= UBound(ChidrenVertices) Then
    
        offset = Point3dSubtract(MasterVertices(index(0)), ChidrenVertices(index(1)))
    
        ' Move child element
        Select Case True
            Case ChidrenEl.IsLineElement
                ChidrenEl.AsLineElement.Move offset
            Case ChidrenEl.IsShapeElement
                ChidrenEl.AsShapeElement.Move offset
        End Select
        
        ChidrenEl.Rewrite
        
        ' Refresh object for refresh sub element
        RefreshoCell CellEl
            
        ' The operation was successful
        MoveElementToVertexInCell = True
        Exit Function
    End If
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.LogError Err.Description, "CellRedreaw.MoveElementToVertexInCell"
    MoveElementToVertexInCell = False
End Function

'Function to rotate a element like another element within a cell.
Private Function RotateElementLikeElementInCell(CellEl As CellElement, ElementToRotateIndex As Integer, ElementWithRotationIndex As Integer, OriginOfRotation As Point3d, Optional AngleOffset As Double) As Boolean
    On Error GoTo ErrorHandler

    Dim ElementToRotate As element
    Dim ElementWithRotation As element
    Dim LineVector As Point3d
    Dim TriangleVector As Point3d
    Dim AngleRad As Double
    Dim RotationMatrix As Matrix3d
    Dim Transform As Transform3d
    Dim Vertices() As Point3d
    
    ' Get the elements from the CellElement using the provided indices
    Set ElementToRotate = GetElementAtIndex(CellEl, ElementToRotateIndex)
    If ElementToRotate Is Nothing Then Exit Function

    Set ElementWithRotation = GetElementAtIndex(CellEl, ElementWithRotationIndex)
    If ElementWithRotation Is Nothing Then Exit Function
    
    ' Determine the direction vector for the line
    LineVector = Point3dSubtract(ElementWithRotation.AsLineElement.EndPoint, ElementWithRotation.AsLineElement.StartPoint)

    ' Determine the direction vector for the triangle
    Vertices = ElementToRotate.AsShapeElement.GetVertices
    TriangleVector = Point3dSubtract(Vertices(1), Vertices(0))

    ' Calculate the signed angle between the vectors
    AngleRad = Point3dSignedAngleBetweenVectors(TriangleVector, LineVector, Point3dFromXYZ(0, 0, 1))

    AngleRad = AngleRad - AngleOffset
    
    ' Create a rotation matrix for the calculated angle
    RotationMatrix = Matrix3dFromVectorAndRotationAngle(Point3dFromXYZ(0, 0, 1), AngleRad)

    ' Create a transform from the rotation matrix and the origin of rotation
    Transform = Transform3dFromMatrix3dAndFixedPoint3d(RotationMatrix, OriginOfRotation)

    ' Apply the transform to the triangle
    ElementToRotate.Transform Transform

    ' Rewrite the triangle to apply the rotation
    ElementToRotate.Rewrite
    
    ' Refresh object for refresh sub element
    RefreshoCell CellEl

    ' The operation was successful
    RotateElementLikeElementInCell = True
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.LogError Err.Description, "CellRedreaw.RotateElementLikeElementInCell"
    RotateElementLikeElementInCell = False
End Function

' Function to calculate the delta (difference) between two coordinates of elements.
Private Function CalculateCoordinateDelta(el1 As element, index1 As Integer, el2 As element, index2 As Integer) As Point3d
    On Error GoTo ErrorHandler
    
    Dim Coordinate1 As Point3d
    Dim Coordinate2 As Point3d
    
    ' Get the coordinates from the elements using the provided indices
    Coordinate1 = GetCoordinate(el1, index1)
    Coordinate2 = GetCoordinate(el2, index2)
    
    ' Calculate the delta between the two coordinates
    CalculateCoordinateDelta = Application.Point3dSubtract(Coordinate1, Coordinate2)
    Exit Function
    
ErrorHandler:
        ' Log the error and return Point3dZero to indicate failure
        ErrorHandler.LogError Err.Description, "CellRedreaw.CalculateCoordinateDelta"
        CalculateCoordinateDelta = Point3dZero
End Function
Private Function MoveVertexToCoordinateInCell(CellEl As CellElement, SubElIndex As Integer, VertexIndex As Integer, Coordinate As Point3d) As Boolean
    On Error GoTo ErrorHandler
    
    Dim subel As element
    
    Set subel = GetElementAtIndex(CellEl, SubElIndex)
    If subel Is Nothing Then Exit Function
    
    Select Case True
        Case subel.IsLineElement
            subel.AsLineElement.ModifyVertex VertexIndex, Coordinate
        Case subel.IsShapeElement
            subel.AsShapeElement.ModifyVertex VertexIndex, Coordinate
        Case Else
            ' Unsupported sub element type
            Exit Function
    End Select
    
    subel.Rewrite
    
    ' Refresh object for refresh sub element
    RefreshoCell CellEl
    
    ' The operation was successful
    MoveVertexToCoordinateInCell = True

    Exit Function
    
ErrorHandler:
    ErrorHandler.LogError Err.Description, "CellRedreaw.MoveVertexToCoordinateInCell"
    MoveVertexToCoordinateInCell = False
End Function

Private Function FindClosestVertex(Point As Point3d, Vertices() As Point3d) As Integer
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim Distance() As Double
    Dim minDistance As Double
    Dim minIndex As Integer
    
    ReDim Distance(UBound(Vertices))
    
    For i = 0 To UBound(Vertices)
        Distance(i) = Point3dDistance(Point, Vertices(i))
    Next i
    
    minDistance = Distance(LBound(Distance))
    minIndex = LBound(Distance)
    
    For i = LBound(Distance) + 1 To UBound(Distance)
        If Distance(i) < minDistance Then
            minDistance = Distance(i)
            minIndex = i
        End If
    Next i
    
    FindClosestVertex = minIndex
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.LogError Err.Description, "CellRedreaw.FindClosestVertex"
    FindClosestVertex = ARESConstants.ARES_CELL_INDEX_ERROR_VALUE
End Function
