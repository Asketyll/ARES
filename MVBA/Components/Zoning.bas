Attribute VB_Name = "Zoning"
' Module: Zoning
' Description: This module provides functions to create zoning (buffer/offset) around elements with arc corners
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ErrorHandlerClass, GetElements, MicroStationDefinition, ARESConstants
Option Explicit

' === PUBLIC FUNCTIONS ===

' Main function to create zoning around elements from specified levels
' Parameters:
'   Lvls() - Array of level names to process
'   Dist - Buffer distance around elements
'   OutputLevel - Name of the level for zoning output (default: "ARES_Zoning")
'   Color - Color for zoning elements (default: 5)
'   Style - Line style for zoning elements (default: 0)
'   Weight - Line weight for zoning elements (default: 1)
Public Sub Zoning(Lvls() As String, _
                  Dist As Double, _
                  Optional OutputLevel As String = "ARES_Zoning", _
                  Optional Color As Long = 5, _
                  Optional Style As Long = 0, _
                  Optional Weight As Long = 1)
    On Error GoTo ErrorHandler

    Dim Elements As Variant
    Dim BufferedElements As Variant
    Dim MergedElements As Variant
    Dim i As Long
    Dim TargetLevel As Level

    ' Validate input parameters
    If Dist <= 0 Then
        ErrorHandler.HandleError "Distance must be greater than zero", 0, "Zoning.Zoning", "ERROR"
        Exit Sub
    End If

    If UBound(Lvls) < LBound(Lvls) Then
        ErrorHandler.HandleError "No levels provided", 0, "Zoning.Zoning", "ERROR"
        Exit Sub
    End If

    ' Check if there is an active model reference
    If Not Application.HasActiveModelReference Then
        ErrorHandler.HandleError "No active model reference", 0, "Zoning.Zoning", "ERROR"
        Exit Sub
    End If

    ' Get or create the output level
    Set TargetLevel = GetOrCreateLevel(OutputLevel, Color, Style, Weight)
    If TargetLevel Is Nothing Then
        ErrorHandler.HandleError "Failed to get or create output level: " & OutputLevel, 0, "Zoning.Zoning", "ERROR"
        Exit Sub
    End If

    ' Get all graphical elements from specified levels (excluding rasters)
    Dim ee As ElementEnumerator
    Dim CurrentElement As Element
    Dim ElementCount As Long

    ' Use GetElements.ByEE to get element enumerator
    Set ee = GetElements.ByEE(Levels:=Lvls)

    ' Build array from enumerator contents
    Elements = ee.BuildArrayFromContents

    ' Check if elements were found
    If IsArray(Elements) Then
        If UBound(Elements) < LBound(Elements) Then
            ErrorHandler.HandleError "No elements found on specified levels", 0, "Zoning.Zoning", "WARNING"
            Exit Sub
        End If
    Else
        ErrorHandler.HandleError "Failed to retrieve elements", 0, "Zoning.Zoning", "ERROR"
        Exit Sub
    End If

    ' Filter out raster elements
    Elements = FilterOutRasterElements(Elements)

    ' Create buffer around each element
    ReDim BufferedElements(LBound(Elements) To UBound(Elements))
    For i = LBound(Elements) To UBound(Elements)
        Set BufferedElements(i) = CreateBufferAroundElement(Elements(i), Dist)
        If BufferedElements(i) Is Nothing Then
            ErrorHandler.HandleError "Failed to create buffer for element ID: " & Elements(i).ID, 0, "Zoning.Zoning", "WARNING"
        End If
    Next i

    ' Merge overlapping zones
    MergedElements = MergeOverlappingZones(BufferedElements)

    ' Apply graphical properties and place on output level
    ApplyZoningProperties MergedElements, TargetLevel, Color, Style, Weight

    ' Refresh the active view
    Dim oView As View
    Set oView = CommandState.LastView
    If Not oView Is Nothing Then
        oView.Redraw
    End If

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, "Zoning.Zoning", "ERROR"
End Sub

' === PRIVATE HELPER FUNCTIONS ===

' Function to get an existing level or create a new one
Private Function GetOrCreateLevel(ByVal LevelName As String, _
                                 ByVal Color As Long, _
                                 ByVal Style As Long, _
                                 ByVal Weight As Long) As Level
    On Error GoTo ErrorHandler

    Dim Level As Level

    Set GetOrCreateLevel = Nothing

    ' Check if level exists
    If Not GetElements.IsValidLevelName(LevelName) Then
        ' Level doesn't exist, create it
        Set Level = ActiveDesignFile.AddNewLevel(LevelName)

        ' Set level properties
        Level.ElementColor = Color
        Level.ElementLineStyle = Style
        Level.ElementLineWeight = Weight

        ' Save changes
        ActiveDesignFile.Levels.Rewrite
    Else
        ' Level exists, get it
        Set Level = ActiveDesignFile.Levels(LevelName)
    End If

    Set GetOrCreateLevel = Level

    Exit Function

ErrorHandler:
    Set GetOrCreateLevel = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, "Zoning.GetOrCreateLevel", "ERROR"
End Function

' Function to create buffer/offset around an element with arc corners
' NOTE: This function requires implementation based on chosen approach:
'       - Option 1: Use MicroStation OFFSET command via CommandState
'       - Option 2: Use Element.ConstructOffset() API if available
'       - Option 3: Manual geometric calculation
Private Function CreateBufferAroundElement(ByVal El As Element, ByVal Distance As Double) As Element
    On Error GoTo ErrorHandler

    Dim BufferedElement As Element
    Dim ElType As MsdElementType

    Set CreateBufferAroundElement = Nothing

    ' Check if element is graphical
    If Not El.IsGraphical Then
        Exit Function
    End If

    ' Get element type
    ElType = El.Type

    ' Handle different element types
    Select Case ElType
        ' === RASTER TYPES (SKIP) ===
        Case msdElementTypeRasterHeader, _
             msdElementTypeRasterComponent, _
             msdElementTypeRasterReference, _
             msdElementTypeRasterReferenceComponent, _
             msdElementTypeRasterFrame
            ' Skip raster elements
            Exit Function

        ' === TEXT ELEMENTS (USE BOUNDING BOX) ===
        Case msdElementTypeText, msdElementTypeTextNode
            Set BufferedElement = CreateBufferFromBoundingBox(El, Distance)

        ' === CELL ELEMENTS (USE CELL HEADER ONLY) ===
        Case msdElementTypeCellHeader, msdElementTypeSharedCell
            Set BufferedElement = CreateBufferFromBoundingBox(El, Distance)

        ' === LINEAR ELEMENTS ===
        Case msdElementTypeLine, _
             msdElementTypeLineString, _
             msdElementTypeShape, _
             msdElementTypeComplexString, _
             msdElementTypeComplexShape
            ' TODO: Implement buffer with arc corners for linear elements
            ' This requires geometric offset calculation
            Set BufferedElement = CreateBufferWithArcs(El, Distance)

        ' === CURVED ELEMENTS ===
        Case msdElementTypeArc, _
             msdElementTypeEllipse, _
             msdElementTypeCurve, _
             msdElementTypeConic, _
             msdElementTypeBsplineCurve
            ' TODO: Implement symmetric buffer for curved elements
            Set BufferedElement = CreateSymmetricCurveBuffer(El, Distance)

        ' === OTHER GRAPHICAL ELEMENTS (USE BOUNDING BOX) ===
        Case Else
            If El.IsGraphical Then
                Set BufferedElement = CreateBufferFromBoundingBox(El, Distance)
            End If
    End Select

    Set CreateBufferAroundElement = BufferedElement

    Exit Function

ErrorHandler:
    Set CreateBufferAroundElement = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, "Zoning.CreateBufferAroundElement", "ERROR"
End Function

' Function to create buffer from element's bounding box
Private Function CreateBufferFromBoundingBox(ByVal El As Element, ByVal Distance As Double) As Element
    On Error GoTo ErrorHandler

    Dim Rng As Range3d
    Dim Points(0 To 4) As Point3d
    Dim ShapeElement As ShapeElement
    Dim i As Integer

    Set CreateBufferFromBoundingBox = Nothing

    ' Get element range
    Rng = El.Range

    ' Create rectangle points with buffer distance
    ' Bottom-left
    Points(0).X = Rng.Low.X - Distance
    Points(0).Y = Rng.Low.Y - Distance
    Points(0).Z = Rng.Low.Z

    ' Bottom-right
    Points(1).X = Rng.High.X + Distance
    Points(1).Y = Rng.Low.Y - Distance
    Points(1).Z = Rng.Low.Z

    ' Top-right
    Points(2).X = Rng.High.X + Distance
    Points(2).Y = Rng.High.Y + Distance
    Points(2).Z = Rng.Low.Z

    ' Top-left
    Points(3).X = Rng.Low.X - Distance
    Points(3).Y = Rng.High.Y + Distance
    Points(3).Z = Rng.Low.Z

    ' Close the shape
    Points(4) = Points(0)

    ' Create shape element
    Set ShapeElement = CreateShapeElement1(Nothing, Points)

    Set CreateBufferFromBoundingBox = ShapeElement

    Exit Function

ErrorHandler:
    Set CreateBufferFromBoundingBox = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, "Zoning.CreateBufferFromBoundingBox", "ERROR"
End Function

' Function to create buffer with arc corners for linear elements
Private Function CreateBufferWithArcs(ByVal El As Element, ByVal Distance As Double) As Element
    On Error GoTo ErrorHandler

    Set CreateBufferWithArcs = Nothing

    Dim ElType As MsdElementType
    ElType = El.Type

    ' For linear elements, create offset using geometric construction
    Select Case ElType
        Case msdElementTypeLine
            Set CreateBufferWithArcs = CreateLineBuffer(El, Distance)

        Case msdElementTypeLineString, _
             msdElementTypeShape
            Set CreateBufferWithArcs = CreatePolylineBuffer(El, Distance)

        Case msdElementTypeComplexString, _
             msdElementTypeComplexShape
            Set CreateBufferWithArcs = CreateComplexBuffer(El, Distance)

        Case Else
            ' Fall back to bounding box for unsupported types
            Set CreateBufferWithArcs = CreateBufferFromBoundingBox(El, Distance)
    End Select

    Exit Function

ErrorHandler:
    Set CreateBufferWithArcs = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, "Zoning.CreateBufferWithArcs", "ERROR"
End Function

' Function to create symmetric buffer for curved elements
Private Function CreateSymmetricCurveBuffer(ByVal El As Element, ByVal Distance As Double) As Element
    On Error GoTo ErrorHandler

    Set CreateSymmetricCurveBuffer = Nothing

    Dim ElType As MsdElementType
    ElType = El.Type

    ' For curved elements, create symmetric offset
    Select Case ElType
        Case msdElementTypeArc
            Set CreateSymmetricCurveBuffer = CreateArcBuffer(El, Distance)

        Case msdElementTypeEllipse
            Set CreateSymmetricCurveBuffer = CreateEllipseBuffer(El, Distance)

        Case msdElementTypeCurve, _
             msdElementTypeConic, _
             msdElementTypeBsplineCurve
            ' Complex curves - use bounding box approach for now
            Set CreateSymmetricCurveBuffer = CreateBufferFromBoundingBox(El, Distance)

        Case Else
            ' Fall back to bounding box
            Set CreateSymmetricCurveBuffer = CreateBufferFromBoundingBox(El, Distance)
    End Select

    Exit Function

ErrorHandler:
    Set CreateSymmetricCurveBuffer = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, "Zoning.CreateSymmetricCurveBuffer", "ERROR"
End Function

' === GEOMETRIC BUFFER CREATION FUNCTIONS ===

' Create buffer around a line element
Private Function CreateLineBuffer(ByVal El As Element, ByVal Distance As Double) As Element
    On Error GoTo ErrorHandler

    Set CreateLineBuffer = Nothing

    Dim LineEl As LineElement
    Dim StartPt As Point3d
    Dim EndPt As Point3d
    Dim Points() As Point3d
    Dim ShapeEl As ShapeElement

    ' Cast to LineElement
    Set LineEl = El.AsLineElement

    ' Get line endpoints
    StartPt = LineEl.StartPoint
    EndPt = LineEl.EndPoint

    ' Calculate perpendicular offset vector
    Dim DX As Double, DY As Double, Length As Double
    Dim OffsetX As Double, OffsetY As Double

    DX = EndPt.X - StartPt.X
    DY = EndPt.Y - StartPt.Y
    Length = Sqr(DX * DX + DY * DY)

    If Length > 0 Then
        ' Perpendicular vector (normalized and scaled by distance)
        OffsetX = -DY / Length * Distance
        OffsetY = DX / Length * Distance

        ' Create rectangle around line with rounded ends
        ReDim Points(0 To 4)

        ' Offset line to one side
        Points(0).X = StartPt.X + OffsetX
        Points(0).Y = StartPt.Y + OffsetY
        Points(0).Z = StartPt.Z

        Points(1).X = EndPt.X + OffsetX
        Points(1).Y = EndPt.Y + OffsetY
        Points(1).Z = EndPt.Z

        ' Offset line to other side
        Points(2).X = EndPt.X - OffsetX
        Points(2).Y = EndPt.Y - OffsetY
        Points(2).Z = EndPt.Z

        Points(3).X = StartPt.X - OffsetX
        Points(3).Y = StartPt.Y - OffsetY
        Points(3).Z = StartPt.Z

        ' Close the shape
        Points(4) = Points(0)

        ' Create shape element
        Set ShapeEl = CreateShapeElement1(Nothing, Points)
        Set CreateLineBuffer = ShapeEl
    Else
        ' Degenerate line - use point buffer (circle)
        Set CreateLineBuffer = CreatePointBuffer(StartPt, Distance)
    End If

    Exit Function

ErrorHandler:
    Set CreateLineBuffer = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, "Zoning.CreateLineBuffer", "ERROR"
End Function

' Create buffer around a point (circle)
Private Function CreatePointBuffer(ByRef Pt As Point3d, ByVal Radius As Double) As Element
    On Error GoTo ErrorHandler

    Set CreatePointBuffer = Nothing

    Dim EllipseEl As EllipseElement
    Dim Rmatrix As Matrix3d

    ' Create identity rotation matrix
    Rmatrix = Matrix3dIdentity()

    ' Create circle as ellipse with equal radii
    Set EllipseEl = CreateEllipseElement2(Nothing, Pt, Radius, Radius, Rmatrix)
    Set CreatePointBuffer = EllipseEl

    Exit Function

ErrorHandler:
    Set CreatePointBuffer = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, "Zoning.CreatePointBuffer", "ERROR"
End Function

' Create buffer around polyline/shape element
Private Function CreatePolylineBuffer(ByVal El As Element, ByVal Distance As Double) As Element
    On Error GoTo ErrorHandler

    Set CreatePolylineBuffer = Nothing

    ' For polylines/shapes, use simplified bounding box approach with buffer
    ' A true offset with arc corners would require complex geometric algorithms
    Set CreatePolylineBuffer = CreateBufferFromBoundingBox(El, Distance)

    ' TODO: Implement true polyline offset with arc corners
    ' This would require:
    ' 1. Extract all vertices
    ' 2. Calculate offset lines for each segment
    ' 3. Calculate intersections or add arcs at corners
    ' 4. Build complex shape from results

    Exit Function

ErrorHandler:
    Set CreatePolylineBuffer = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, "Zoning.CreatePolylineBuffer", "ERROR"
End Function

' Create buffer around complex element
Private Function CreateComplexBuffer(ByVal El As Element, ByVal Distance As Double) As Element
    On Error GoTo ErrorHandler

    Set CreateComplexBuffer = Nothing

    ' For complex elements, use bounding box approach
    Set CreateComplexBuffer = CreateBufferFromBoundingBox(El, Distance)

    ' TODO: Implement by processing sub-elements individually
    ' and combining the results

    Exit Function

ErrorHandler:
    Set CreateComplexBuffer = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, "Zoning.CreateComplexBuffer", "ERROR"
End Function

' Create buffer around arc element
Private Function CreateArcBuffer(ByVal El As Element, ByVal Distance As Double) As Element
    On Error GoTo ErrorHandler

    Set CreateArcBuffer = Nothing

    Dim ArcEl As ArcElement
    Dim Origin As Point3d
    Dim PrimaryRadius As Double
    Dim SecondaryRadius As Double
    Dim Rmatrix As Matrix3d
    Dim StartAngle As Double
    Dim SweepAngle As Double
    Dim NewArcEl As ArcElement

    ' Cast to ArcElement
    Set ArcEl = El.AsArcElement

    ' Get arc properties
    Origin = ArcEl.Origin
    PrimaryRadius = ArcEl.PrimaryRadius
    SecondaryRadius = ArcEl.SecondaryRadius
    Rmatrix = ArcEl.Rmatrix
    StartAngle = ArcEl.StartAngle
    SweepAngle = ArcEl.SweepAngle

    ' Create offset arc with increased radius
    Set NewArcEl = CreateArcElement3(Nothing, Origin, PrimaryRadius + Distance, _
                                     SecondaryRadius + Distance, Rmatrix, _
                                     StartAngle, SweepAngle)
    Set CreateArcBuffer = NewArcEl

    ' TODO: For full buffer, create shape with inner and outer arcs

    Exit Function

ErrorHandler:
    Set CreateArcBuffer = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, "Zoning.CreateArcBuffer", "ERROR"
End Function

' Create buffer around ellipse element
Private Function CreateEllipseBuffer(ByVal El As Element, ByVal Distance As Double) As Element
    On Error GoTo ErrorHandler

    Set CreateEllipseBuffer = Nothing

    Dim EllipseEl As EllipseElement
    Dim Origin As Point3d
    Dim PrimaryRadius As Double
    Dim SecondaryRadius As Double
    Dim Rmatrix As Matrix3d
    Dim NewEllipseEl As EllipseElement

    ' Cast to EllipseElement
    Set EllipseEl = El.AsEllipseElement

    ' Get ellipse properties
    Origin = EllipseEl.Origin
    PrimaryRadius = EllipseEl.PrimaryRadius
    SecondaryRadius = EllipseEl.SecondaryRadius
    Rmatrix = EllipseEl.Rmatrix

    ' Create offset ellipse with increased radii
    Set NewEllipseEl = CreateEllipseElement2(Nothing, Origin, _
                                             PrimaryRadius + Distance, _
                                             SecondaryRadius + Distance, _
                                             Rmatrix)
    Set CreateEllipseBuffer = NewEllipseEl

    Exit Function

ErrorHandler:
    Set CreateEllipseBuffer = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, "Zoning.CreateEllipseBuffer", "ERROR"
End Function

' Function to merge overlapping zones
' TODO: Implement zone merging algorithm
Private Function MergeOverlappingZones(BufferedElements As Variant) As Variant
    On Error GoTo ErrorHandler

    ' PLACEHOLDER: This requires implementation of geometric union algorithm
    ' to merge overlapping zones
    '
    ' Possible approaches:
    ' 1. Use MicroStation API for Boolean operations if available
    ' 2. Use external library for geometric unions
    ' 3. Simplified approach: keep all zones without merging
    '
    ' For now, return the buffered elements as-is
    MergeOverlappingZones = BufferedElements

    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, "Zoning.MergeOverlappingZones", "ERROR"
    MergeOverlappingZones = BufferedElements
End Function

' Function to filter out raster elements from an array
Private Function FilterOutRasterElements(Elements As Variant) As Variant
    On Error GoTo ErrorHandler

    Dim FilteredElements() As Element
    Dim Count As Long
    Dim i As Long
    Dim El As Element

    Count = 0

    ' Check if elements is an array
    If Not IsArray(Elements) Then
        FilterOutRasterElements = Elements
        Exit Function
    End If

    ' Count non-raster elements
    For i = LBound(Elements) To UBound(Elements)
        If Not Elements(i) Is Nothing Then
            Set El = Elements(i)
            If Not IsRasterElement(El) Then
                Count = Count + 1
            End If
        End If
    Next i

    ' Build filtered array
    If Count > 0 Then
        ReDim FilteredElements(1 To Count)
        Count = 0
        For i = LBound(Elements) To UBound(Elements)
            If Not Elements(i) Is Nothing Then
                Set El = Elements(i)
                If Not IsRasterElement(El) Then
                    Count = Count + 1
                    Set FilteredElements(Count) = El
                End If
            End If
        Next i
    Else
        ReDim FilteredElements(0)
    End If

    FilterOutRasterElements = FilteredElements
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, "Zoning.FilterOutRasterElements", "ERROR"
    FilterOutRasterElements = Elements
End Function

' Function to check if an element is a raster type
Private Function IsRasterElement(ByVal El As Element) As Boolean
    On Error GoTo ErrorHandler

    Dim ElType As MsdElementType
    ElType = El.Type

    Select Case ElType
        Case msdElementTypeRasterHeader, _
             msdElementTypeRasterComponent, _
             msdElementTypeRasterReference, _
             msdElementTypeRasterReferenceComponent, _
             msdElementTypeRasterFrame
            IsRasterElement = True
        Case Else
            IsRasterElement = False
    End Select

    Exit Function

ErrorHandler:
    IsRasterElement = False
    ErrorHandler.HandleError Err.Description, Err.Number, "Zoning.IsRasterElement", "ERROR"
End Function

' Function to apply graphical properties and place elements on target level
Private Sub ApplyZoningProperties(Elements As Variant, _
                                 ByRef TargetLevel As Level, _
                                 ByVal Color As Long, _
                                 ByVal Style As Long, _
                                 ByVal Weight As Long)
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim El As Element

    ' Check if elements is an array
    If Not IsArray(Elements) Then Exit Sub

    ' Apply properties to each element
    For i = LBound(Elements) To UBound(Elements)
        If Not Elements(i) Is Nothing Then
            Set El = Elements(i)

            ' Set level
            El.Level = TargetLevel

            ' Set graphical properties
            El.Color = Color
            El.Style = Style
            El.Weight = Weight

            ' Add to model
            ActiveModelReference.AddElement El
        End If
    Next i

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, "Zoning.ApplyZoningProperties", "ERROR"
End Sub
