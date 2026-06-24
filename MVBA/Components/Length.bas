' Module: Length
' Description: This module provides functions to calculate lengths of elements in MicroStation with silent error handling.
' The module includes functions to determine the length of various element types, handle rounding logic,
' NEVER USE Rnd = 255 in GetLength function! It is reserved for errors.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: Config, ARESConfigClass, ARESConstants, LangManager, ErrorHandlerClass
Option Explicit

' Public function to get the length of an element
Public Function GetLength(ByVal El As element, Optional RND As Variant, Optional RndLength As Boolean = True, Optional ErasRnd As Boolean = False) As Double
    On Error GoTo ErrorHandler
    ' Determine the length based on the element type
    GetLength = GetElementLength(El)
    ' Handle rounding if required
    If RndLength Then
        RND = HandleRounding(RND, ErasRnd)
        If RND = ARES_RND_ERROR_VALUE Then
            ShowStatus GetTranslation("LengthRoundError") & ARES_RND_ERROR_VALUE
            GetLength = 0
            Exit Function
        End If
        GetLength = RoundedLength(GetLength, CByte(RND))
    ElseIf ErasRnd Then
        If Not HandleRoundingForErase(RND) Then
            GetLength = 0
            Exit Function
        End If
    End If
    Exit Function

ErrorHandler:
    ' Return 0 in case of an error
    GetLength = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.GetLength"
End Function

' Private function to get the length of an element based on its type
Private Function GetElementLength(ByVal El As element) As Double
    On Error GoTo ErrorHandler
    ' Determine the length based on the element type
    Select Case True
        Case El.IsComplexStringElement
            GetElementLength = El.AsComplexStringElement.Length
        Case El.IsComplexShapeElement
            GetElementLength = LengthComplexShape(El, True)
        Case El.IsLineElement
            GetElementLength = El.AsLineElement.Length
        Case El.IsArcElement
            GetElementLength = El.AsArcElement.Length
        Case El.IsShapeElement
            GetElementLength = LengthShape(El, True)
        Case Else
            GetElementLength = 0
            ShowStatus GetTranslation("LengthElementTypeNotSupportedByInterface", DLongToString(El.ID), El.Type)
    End Select
    
    Exit Function

ErrorHandler:
    ' Return 0 in case of an error
    GetElementLength = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.GetElementLength"
End Function

' Private function to handle rounding logic
Private Function HandleRounding(Optional RND As Variant, Optional ErasRnd As Boolean) As Variant
    On Error GoTo ErrorHandler
    ' Handle missing rounding value
    If IsMissing(RND) Then
        RND = GetRoundValue()
    ElseIf ErasRnd And (VarType(RND) = vbByte Or VarType(RND) = vbInteger) Then
        ' Set rounding value if erase rounding is true
        If Not SetRound(CByte(RND)) Then
            HandleRounding = ARES_RND_ERROR_VALUE
            Exit Function
        End If
    End If
    HandleRounding = RND
    Exit Function

ErrorHandler:
    ' Return error value in case of an error
    HandleRounding = ARES_RND_ERROR_VALUE
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.HandleRounding"
End Function

' Private function to handle rounding logic for erase
Private Function HandleRoundingForErase(Optional RND As Variant) As Boolean
    On Error GoTo ErrorHandler
    HandleRoundingForErase = False
    ' Set default rounding if Rnd is missing
    If IsMissing(RND) Then
        If Not ResetRound() Then Exit Function
    ElseIf VarType(RND) = vbByte Or VarType(RND) = vbInteger Then
        ' Set rounding value if Rnd is provided
        If Not SetRound(CByte(RND)) Then Exit Function
    End If
    HandleRoundingForErase = True
    Exit Function

ErrorHandler:
    ' Return False in case of an error
    HandleRoundingForErase = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.HandleRoundingForErase"
End Function

' Private function to calculate the length of a complex shape element
Private Function LengthComplexShape(ByVal El As ComplexShapeElement, Optional ByVal LongestSideOnly As Boolean = False) As Double
    On Error GoTo ErrorHandler
    
    If LongestSideOnly Then
        ' Find the longest sub-element
        LengthComplexShape = GetLongestSideFromComplexShape(El)
    Else
        ' Return the perimeter (default behavior)
        LengthComplexShape = El.Perimeter
    End If
    
    Exit Function

ErrorHandler:
    ' Return 0 in case of an error
    LengthComplexShape = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.LengthComplexShape"
End Function

' Private function to calculate the length of a shape element
Private Function LengthShape(ByVal El As ShapeElement, Optional ByVal LongestSideOnly As Boolean = False) As Double
    On Error GoTo ErrorHandler
    
    If LongestSideOnly Then
        ' Calculate the length of the longest side
        LengthShape = GetLongestSideFromShape(El)
    Else
        ' Return the perimeter (default behavior)
        LengthShape = El.Perimeter
    End If
    
    Exit Function

ErrorHandler:
    ' Return 0 in case of an error
    LengthShape = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.LengthShape"
End Function

' Helper function to get the longest side from a simple shape
Private Function GetLongestSideFromShape(ByVal El As ShapeElement) As Double
    On Error GoTo ErrorHandler
    
    Dim Vertices() As Point3d
    Dim i As Long
    Dim SideLength As Double
    Dim LongestSide As Double
    
    LongestSide = 0
    Vertices = El.GetVertices()
    
    ' Calculate length of each side and find the longest
    For i = 0 To UBound(Vertices) - 1
        SideLength = Point3dDistance(Vertices(i), Vertices(i + 1))
        If SideLength > LongestSide Then
            LongestSide = SideLength
        End If
    Next i
    
    GetLongestSideFromShape = LongestSide
    Exit Function

ErrorHandler:
    GetLongestSideFromShape = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.GetLongestSideFromShape"
End Function

' Helper function to get the longest side from a complex shape
Private Function GetLongestSideFromComplexShape(ByVal El As ComplexShapeElement, Optional ByVal Depth As Integer = 0) As Double
    On Error GoTo ErrorHandler

    Const MAX_DEPTH As Integer = 10
    If Depth > MAX_DEPTH Then
        GetLongestSideFromComplexShape = 0
        Exit Function
    End If

    Dim ELEnum As ElementEnumerator
    Dim subel As element
    Dim ElementLength As Double
    Dim LongestSide As Double

    LongestSide = 0
    Set ELEnum = El.GetSubElements

    ' Iterate through sub-elements and find the longest one
    Do While ELEnum.MoveNext
        Set subel = ELEnum.Current
        ElementLength = 0

        Select Case True
            Case subel.IsLineElement
                ElementLength = subel.AsLineElement.Length
            Case subel.IsArcElement
                ElementLength = subel.AsArcElement.Length
            Case subel.IsShapeElement
                ' For nested shapes, get their longest side recursively
                ElementLength = GetLongestSideFromShape(subel.AsShapeElement)
            Case subel.IsComplexShapeElement
                ' For nested complex shapes, get their longest side recursively
                ElementLength = GetLongestSideFromComplexShape(subel.AsComplexShapeElement, Depth + 1)
        End Select

        If ElementLength > LongestSide Then
            LongestSide = ElementLength
        End If
    Loop

    GetLongestSideFromComplexShape = LongestSide
    Exit Function

ErrorHandler:
    GetLongestSideFromComplexShape = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.GetLongestSideFromComplexShape"
End Function

' Private function to round the length to a specified number of decimal places
Private Function RoundedLength(Length As Double, RND As Byte) As Double
    On Error GoTo ErrorHandler
    RoundedLength = Round(Length, RND)
    Exit Function

ErrorHandler:
    ' Return 0 in case of an error
    RoundedLength = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.RoundedLength"
End Function

' Private function to get the rounding value
Private Function GetRoundValue() As Variant
    On Error GoTo ErrorHandler
    Dim roundValue As String
    roundValue = ARESConfig.ARES_ROUNDS.Value
    ' Handle empty rounding value - reset to default and read it back
    If roundValue = "" Then
        If ResetRound() Then
            roundValue = ARESConfig.ARES_ROUNDS.defaultValue
        Else
            GetRoundValue = ARES_RND_ERROR_VALUE
            Exit Function
        End If
    End If
    GetRoundValue = CByte(roundValue)
    Exit Function

ErrorHandler:
    ' Return error value in case of an error
    GetRoundValue = ARES_RND_ERROR_VALUE
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.GetRoundValue"
End Function

' Public function to set the rounding configuration variable
Public Function SetRound(RND As Byte) As Boolean
    On Error GoTo ErrorHandler
    If RND <> ARES_RND_ERROR_VALUE Then
        SetRound = Config.SetVar(ARESConfig.ARES_ROUNDS.key, RND)
    Else
        ShowStatus GetTranslation("LengthRoundError") & ARES_RND_ERROR_VALUE
    End If
    Exit Function

ErrorHandler:
    ' Return False in case of an error
    SetRound = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.SetRound"
End Function

' Public function to reset the rounding configuration variable
Public Function ResetRound() As Boolean
    On Error GoTo ErrorHandler
    ARESConfig.ResetConfigVar ARESConfig.ARES_ROUNDS.key
    ResetRound = (ARESConfig.ARES_ROUNDS.Value = ARESConfig.ARES_ROUNDS.defaultValue)
    Exit Function

ErrorHandler:
    ' Return False in case of an error
    ResetRound = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.ResetRound"
End Function

' ============================================================
'  PARTIAL LENGTH (zone-aware)
' ============================================================

' GetPartialLengthInsideZones
' Single entry point. Calls CollectIntersectionPoints once at the element level.
'   nPts = 0  — fast path: StartPoint for open types; first-vertex + zone-center for closed.
'   nPts > 0  — routes to typed helpers, passing pre-collected intersection points for
'               Line/Arc; sub-element decomposition via PartialLengthPrimitive for others.
Public Function GetPartialLengthInsideZones(ByVal oEl As Element, _
                                             ByRef oZones() As Element) As Double
    On Error GoTo ErrorHandler

    If Not oEl.IsIntersectableElement Then Exit Function

    Dim isectPts() As Point3d
    Dim nPts       As Long
    CollectIntersectionPoints oEl, oZones, isectPts, nPts

    If nPts = 0 Then
        Dim bInside  As Boolean
        Dim oVL      As VertexList
        Dim verts()  As Point3d
        Dim oCSEl    As ComplexShapeElement
        Dim oSubEE   As ElementEnumerator
        Dim i        As Long
        Select Case oEl.Type
            Case msdElementTypeLine
                bInside = PointInAnyZone(oEl.AsLineElement.StartPoint, oZones)
            Case msdElementTypeArc
                bInside = PointInAnyZone(oEl.AsArcElement.StartPoint, oZones)
            Case msdElementTypeLineString, msdElementTypeShape
                Set oVL = oEl
                verts   = oVL.GetVertices
                bInside = PointInAnyZone(verts(LBound(verts)), oZones)
            Case msdElementTypeComplexString
                bInside = PointInAnyZone(oEl.AsComplexStringElement.StartPoint, oZones)
            Case msdElementTypeComplexShape
                Set oCSEl  = oEl.AsComplexShapeElement
                Set oSubEE = oCSEl.GetSubElements
                If oSubEE.MoveNext Then
                    bInside = PointInAnyZone(oSubEE.Current.AsChainableElement.StartPoint, oZones)
                End If
        End Select

        If Not bInside Then Exit Function

        Select Case oEl.Type
            Case msdElementTypeLine
                GetPartialLengthInsideZones = oEl.AsLineElement.Length
            Case msdElementTypeArc
                GetPartialLengthInsideZones = oEl.AsArcElement.Length
            Case msdElementTypeComplexString
                GetPartialLengthInsideZones = oEl.AsComplexStringElement.Length
            Case msdElementTypeLineString
                Set oVL = oEl
                verts   = oVL.GetVertices
                For i = LBound(verts) To UBound(verts) - 1
                    GetPartialLengthInsideZones = GetPartialLengthInsideZones + _
                                                  Point3dDistance(verts(i), verts(i + 1))
                Next i
            Case msdElementTypeShape
                GetPartialLengthInsideZones = LengthShape(oEl.AsShapeElement, False)
            Case msdElementTypeComplexShape
                GetPartialLengthInsideZones = LengthComplexShape(oEl.AsComplexShapeElement, False)
        End Select
        Exit Function
    End If

    ' nPts > 0: boundary crossing — route to typed partial-length helpers.
    Select Case oEl.Type
        Case msdElementTypeLine
            GetPartialLengthInsideZones = PartialLengthLine(oEl.AsLineElement, isectPts, nPts, oZones)
        Case msdElementTypeArc
            GetPartialLengthInsideZones = PartialLengthArc(oEl.AsArcElement, isectPts, nPts, oZones)
        Case msdElementTypeLineString
            GetPartialLengthInsideZones = PartialLengthLineString(oEl, oZones)
        Case msdElementTypeComplexString
            GetPartialLengthInsideZones = PartialLengthComplexString(oEl, oZones)
        Case msdElementTypeShape
            GetPartialLengthInsideZones = PartialLengthShape(oEl, oZones)
        Case msdElementTypeComplexShape
            GetPartialLengthInsideZones = PartialLengthComplexShape(oEl, oZones)
    End Select
    Exit Function
ErrorHandler:
    GetPartialLengthInsideZones = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.GetPartialLengthInsideZones"
End Function

' PointInZone
' Tests whether a point lies strictly inside a closed planar zone using a
' 2D ray-cast: build a horizontal line that extends well past the zone's
' bbox, then count the boundary crossings via GetIntersectionPoints.
' Odd count → inside; even count → outside.
Private Function PointInZone(ByRef pt As Point3d, ByVal oZone As Element) As Boolean
    On Error GoTo ErrorHandler

    Dim oRange     As Range3d
    Dim ptEnd      As Point3d
    Dim dRayLen    As Double
    Dim oRay       As LineElement
    Dim isectPts() As Point3d
    Dim nIsect     As Long

    oRange = oZone.Range

    dRayLen = (oRange.High.X - oRange.Low.X) + Abs(pt.X - oRange.Low.X) + 1#
    ptEnd.X = pt.X + dRayLen
    ptEnd.Y = pt.Y
    ptEnd.Z = pt.Z
    Set oRay = CreateLineElement2(Nothing, pt, ptEnd)

    Dim oRayIsect As IntersectableElement
    Set oRayIsect = oRay
    On Error Resume Next
    isectPts = oRayIsect.GetIntersectionPoints(oZone, Matrix3dIdentity)
    nIsect = -1
    nIsect = UBound(isectPts)
    On Error GoTo ErrorHandler

    If nIsect < 0 Then
        PointInZone = False
        Exit Function
    End If

    PointInZone = (((nIsect + 1) Mod 2) = 1)
    Exit Function

ErrorHandler:
    PointInZone = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.PointInZone"
End Function

' PointInAnyZone
' Returns True if pt lies inside at least one zone (delegates to PointInZone ray-cast).
Private Function PointInAnyZone(ByRef pt As Point3d, ByRef zones() As Element) As Boolean
    On Error GoTo ErrorHandler
    Dim i As Long
    For i = LBound(zones) To UBound(zones)
        If PointInZone(pt, zones(i)) Then
            PointInAnyZone = True
            Exit Function
        End If
    Next i
    PointInAnyZone = False
    Exit Function
ErrorHandler:
    PointInAnyZone = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.PointInAnyZone"
End Function

' AnyZoneCenterInsideElement
' Returns True if the bbox-center of any zone lies inside the closed element oEl.
' Uses the zone's axis-aligned bbox midpoint as a proxy for its centroid — reliable for
' convex/rectangular/elliptical zones. Reuses PointInZone ray-cast with oEl as the target.
Private Function AnyZoneCenterInsideElement(ByRef zones() As Element, _
                                             ByVal oEl As Element) As Boolean
    On Error GoTo ErrorHandler
    Dim i       As Long
    Dim zr      As Range3d
    Dim zCenter As Point3d
    For i = LBound(zones) To UBound(zones)
        zr = zones(i).Range
        zCenter.X = (zr.Low.X + zr.High.X) / 2#
        zCenter.Y = (zr.Low.Y + zr.High.Y) / 2#
        zCenter.Z = (zr.Low.Z + zr.High.Z) / 2#
        If PointInZone(zCenter, oEl) Then
            AnyZoneCenterInsideElement = True
            Exit Function
        End If
    Next i
    AnyZoneCenterInsideElement = False
    Exit Function
ErrorHandler:
    AnyZoneCenterInsideElement = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.AnyZoneCenterInsideElement"
End Function

' CollectIntersectionPoints
' Accumulates every point where oEl crosses the boundary of any zone.
' Near-duplicate points (zone-boundary vertex hits) are removed after collection.
Private Sub CollectIntersectionPoints(ByVal oEl As Element, _
                                       ByRef zones() As Element, _
                                       ByRef outPts() As Point3d, _
                                       ByRef outCount As Long)
    On Error GoTo ErrorHandler
    outCount = 0
    If Not oEl.IsIntersectableElement Then Exit Sub

    Dim oIsect As IntersectableElement
    Set oIsect = oEl.AsIntersectableElement

    Dim i     As Long
    Dim pts() As Point3d
    Dim nPts  As Long
    Dim j     As Long

    For i = LBound(zones) To UBound(zones)
        On Error Resume Next
        pts  = oIsect.GetIntersectionPoints(zones(i), Matrix3dIdentity)
        nPts = -1
        nPts = UBound(pts)
        On Error GoTo ErrorHandler
        If nPts >= 0 Then
            For j = 0 To nPts
                ReDim Preserve outPts(0 To outCount)
                outPts(outCount) = pts(j)
                outCount = outCount + 1
            Next j
        End If
    Next i

    If outCount > 1 Then DeduplicatePoints outPts, outCount
    Exit Sub
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.CollectIntersectionPoints"
End Sub

' DeduplicatePoints
' Removes points within EPSILON of an earlier point. O(n²), n < ~20 in practice.
Private Sub DeduplicatePoints(ByRef pts() As Point3d, ByRef nPts As Long)
    Const EPSILON As Double = 1E-9
    If nPts <= 1 Then Exit Sub

    Dim keep()  As Boolean
    Dim i       As Long
    Dim j       As Long
    Dim iWrite  As Long
    ReDim keep(0 To nPts - 1)
    For i = 0 To nPts - 1: keep(i) = True: Next i

    For i = 0 To nPts - 2
        If keep(i) Then
            For j = i + 1 To nPts - 1
                If keep(j) Then
                    If Point3dDistance(pts(i), pts(j)) < EPSILON Then keep(j) = False
                End If
            Next j
        End If
    Next i

    iWrite = 0
    For i = 0 To nPts - 1
        If keep(i) Then
            pts(iWrite) = pts(i)
            iWrite = iWrite + 1
        End If
    Next i
    nPts = iWrite
End Sub

' SortByDistanceFrom
' Sorts pts(0..nPts-1) by ascending distance from origin. Bubble sort (n is small).
Private Sub SortByDistanceFrom(ByRef pts() As Point3d, ByVal nPts As Long, _
                                ByRef origin As Point3d)
    Dim i   As Long
    Dim j   As Long
    Dim tmp As Point3d
    For i = 0 To nPts - 2
        For j = 0 To nPts - 2 - i
            If Point3dDistance(origin, pts(j)) > Point3dDistance(origin, pts(j + 1)) Then
                tmp        = pts(j)
                pts(j)     = pts(j + 1)
                pts(j + 1) = tmp
            End If
        Next j
    Next i
End Sub

' ArcAngleParam
' Returns the angular offset of pt from the arc's start angle, in [0, |sweepAngle|].
' Used to sort boundary crossings and compute partial arc-length fractions.
Private Function ArcAngleParam(ByRef pt As Point3d, ByVal oArc As ArcElement) As Double
    Dim angle  As Double
    Dim offset As Double
    Dim twoPi  As Double
    twoPi  = 2# * Application.PI
    angle  = Point3dPolarAngle(Point3dSubtract(pt, oArc.CenterPoint))
    offset = angle - oArc.StartAngle
    If oArc.SweepAngle >= 0 Then
        Do While offset < 0:       offset = offset + twoPi: Loop
        Do While offset >= twoPi:  offset = offset - twoPi: Loop
    Else
        Do While offset > 0:       offset = offset - twoPi: Loop
        Do While offset <= -twoPi: offset = offset + twoPi: Loop
    End If
    ArcAngleParam = Abs(offset)
End Function

' SortArcPoints
' Sorts pts(0..nPts-1) by ascending ArcAngleParam. Bubble sort (n is small).
Private Sub SortArcPoints(ByRef pts() As Point3d, ByVal nPts As Long, _
                           ByVal oArc As ArcElement)
    Dim i   As Long
    Dim j   As Long
    Dim tmp As Point3d
    For i = 0 To nPts - 2
        For j = 0 To nPts - 2 - i
            If ArcAngleParam(pts(j), oArc) > ArcAngleParam(pts(j + 1), oArc) Then
                tmp        = pts(j)
                pts(j)     = pts(j + 1)
                pts(j + 1) = tmp
            End If
        Next j
    Next i
End Sub

' PartialLengthLine
' Alternating inside/outside walk with pre-collected intersection points (nPts guaranteed > 0).
Private Function PartialLengthLine(ByVal oLine As LineElement, _
                                    ByRef isectPts() As Point3d, _
                                    ByVal nPts As Long, _
                                    ByRef zones() As Element) As Double
    On Error GoTo ErrorHandler
    Dim startPt As Point3d: startPt = oLine.StartPoint
    Dim endPt   As Point3d: endPt   = oLine.EndPoint
    Dim bInside As Boolean: bInside = PointInAnyZone(startPt, zones)
    Dim prevPt  As Point3d
    Dim dTotal  As Double
    Dim i       As Long
    SortByDistanceFrom isectPts, nPts, startPt
    prevPt = startPt
    For i = 0 To nPts - 1
        If bInside Then dTotal = dTotal + Point3dDistance(prevPt, isectPts(i))
        prevPt  = isectPts(i)
        bInside = Not bInside
    Next i
    If bInside Then dTotal = dTotal + Point3dDistance(prevPt, endPt)
    PartialLengthLine = dTotal
    Exit Function
ErrorHandler:
    PartialLengthLine = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.PartialLengthLine"
End Function

' PartialLengthArc
' Alternating inside/outside walk. Uses ArcSegmentLength (PartialDelete) for each inside
' segment — correct for both circular and elliptical arcs.
Private Function PartialLengthArc(ByVal oArc As ArcElement, _
                                   ByRef isectPts() As Point3d, _
                                   ByVal nPts As Long, _
                                   ByRef zones() As Element) As Double
    On Error GoTo ErrorHandler
    Dim bInside As Boolean: bInside = PointInAnyZone(oArc.StartPoint, zones)
    Dim prevPt  As Point3d: prevPt  = oArc.StartPoint
    Dim dTotal  As Double
    Dim i       As Long
    SortArcPoints isectPts, nPts, oArc
    For i = 0 To nPts - 1
        If bInside Then dTotal = dTotal + ArcSegmentLength(oArc, prevPt, isectPts(i))
        prevPt  = isectPts(i)
        bInside = Not bInside
    Next i
    If bInside Then dTotal = dTotal + ArcSegmentLength(oArc, prevPt, oArc.EndPoint)
    PartialLengthArc = dTotal
    Exit Function
ErrorHandler:
    PartialLengthArc = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.PartialLengthArc"
End Function

' ArcSegmentLength
' Returns the arc length between pt1 and pt2 (both projected onto oArc).
' PartialDelete removes [pt1, pt2] from oArc (non-destructive on the original) and returns
' the outer fragments; inside length = total - outer.
' ViewSpecifier 1 = first view, valid for non-interactive use per MVBA docs.
Private Function ArcSegmentLength(ByVal oArc As ArcElement, _
                                   ByRef pt1 As Point3d, _
                                   ByRef pt2 As Point3d) As Double
    On Error GoTo ErrorHandler
    Dim oPart1   As Element
    Dim oPart2   As Element
    Dim dOutside As Double
    oArc.PartialDelete oPart1, oPart2, pt1, pt2, pt1, 1
    If Not oPart1 Is Nothing Then dOutside = dOutside + oPart1.AsArcElement.Length
    If Not oPart2 Is Nothing Then dOutside = dOutside + oPart2.AsArcElement.Length
    ArcSegmentLength = oArc.Length - dOutside
    Exit Function
ErrorHandler:
    ArcSegmentLength = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.ArcSegmentLength"
End Function

' PartialLengthPrimitive
' Computes partial length for any primitive sub-element (Line, Arc, LineString).
' Collects its own intersection points — used by decomposition loops in
' PartialLengthLineString, PartialLengthShape, PartialLengthComplexString/Shape.
Private Function PartialLengthPrimitive(ByVal oEl As Element, ByRef zones() As Element) As Double
    On Error GoTo ErrorHandler
    If Not oEl.IsIntersectableElement Then Exit Function
    Dim pts() As Point3d
    Dim n     As Long
    CollectIntersectionPoints oEl, zones, pts, n
    If n = 0 Then
        Select Case oEl.Type
            Case msdElementTypeLine
                If PointInAnyZone(oEl.AsLineElement.StartPoint, zones) Then
                    PartialLengthPrimitive = Point3dDistance(oEl.AsLineElement.StartPoint, _
                                                             oEl.AsLineElement.EndPoint)
                End If
            Case msdElementTypeArc
                If PointInAnyZone(oEl.AsArcElement.StartPoint, zones) Then
                    PartialLengthPrimitive = oEl.AsArcElement.Length
                End If
            Case msdElementTypeLineString
                Dim oVL   As VertexList
                Set oVL   = oEl
                Dim verts() As Point3d
                verts     = oVL.GetVertices
                If PointInAnyZone(verts(LBound(verts)), zones) Then
                    Dim j As Long
                    For j = LBound(verts) To UBound(verts) - 1
                        PartialLengthPrimitive = PartialLengthPrimitive + _
                                                  Point3dDistance(verts(j), verts(j + 1))
                    Next j
                End If
        End Select
    Else
        Select Case oEl.Type
            Case msdElementTypeLine
                PartialLengthPrimitive = PartialLengthLine(oEl.AsLineElement, pts, n, zones)
            Case msdElementTypeArc
                PartialLengthPrimitive = PartialLengthArc(oEl.AsArcElement, pts, n, zones)
            Case msdElementTypeLineString
                PartialLengthPrimitive = PartialLengthLineString(oEl, zones)
        End Select
    End If
    Exit Function
ErrorHandler:
    PartialLengthPrimitive = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.PartialLengthPrimitive"
End Function

' PartialLengthLineString
' Decomposes into line segments (CreateLineElement2 orphans); delegates to PartialLengthPrimitive.
Private Function PartialLengthLineString(ByVal oEl As Element, ByRef zones() As Element) As Double
    On Error GoTo ErrorHandler
    Dim oVL        As VertexList
    Set oVL        = oEl
    Dim vertices() As Point3d
    vertices       = oVL.GetVertices
    If UBound(vertices) - LBound(vertices) < 1 Then Exit Function
    Dim dTotal As Double
    Dim i      As Long
    Dim oSeg   As LineElement
    For i = LBound(vertices) To UBound(vertices) - 1
        Set oSeg = CreateLineElement2(Nothing, vertices(i), vertices(i + 1))
        dTotal   = dTotal + PartialLengthPrimitive(oSeg, zones)
    Next i
    PartialLengthLineString = dTotal
    Exit Function
ErrorHandler:
    PartialLengthLineString = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.PartialLengthLineString"
End Function

' PartialLengthShape
' Decomposes into polygon edges (including closing segment); delegates to PartialLengthPrimitive.
Private Function PartialLengthShape(ByVal oEl As Element, ByRef zones() As Element) As Double
    On Error GoTo ErrorHandler
    Dim oVL        As VertexList
    Set oVL        = oEl
    Dim vertices() As Point3d
    vertices       = oVL.GetVertices
    If UBound(vertices) - LBound(vertices) < 1 Then Exit Function
    Dim dTotal As Double
    Dim i      As Long
    Dim oSeg   As LineElement
    For i = LBound(vertices) To UBound(vertices) - 1
        Set oSeg = CreateLineElement2(Nothing, vertices(i), vertices(i + 1))
        dTotal   = dTotal + PartialLengthPrimitive(oSeg, zones)
    Next i
    Set oSeg = CreateLineElement2(Nothing, vertices(UBound(vertices)), vertices(LBound(vertices)))
    dTotal   = dTotal + PartialLengthPrimitive(oSeg, zones)
    PartialLengthShape = dTotal
    Exit Function
ErrorHandler:
    PartialLengthShape = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.PartialLengthShape"
End Function

' PartialLengthComplexString / PartialLengthComplexShape
' Iterate sub-elements via GetSubElements; delegate each to PartialLengthPrimitive.
Private Function PartialLengthComplexString(ByVal oEl As Element, ByRef zones() As Element) As Double
    On Error GoTo ErrorHandler
    Dim cxEl    As ComplexElement
    Set cxEl    = oEl
    Dim subEnum As ElementEnumerator
    Set subEnum = cxEl.GetSubElements()
    Dim dTotal  As Double
    Dim oSub    As Element
    Do While subEnum.MoveNext
        Set oSub = subEnum.Current
        dTotal   = dTotal + PartialLengthPrimitive(oSub, zones)
    Loop
    PartialLengthComplexString = dTotal
    Exit Function
ErrorHandler:
    PartialLengthComplexString = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.PartialLengthComplexString"
End Function

Private Function PartialLengthComplexShape(ByVal oEl As Element, ByRef zones() As Element) As Double
    On Error GoTo ErrorHandler
    Dim cxEl    As ComplexElement
    Set cxEl    = oEl
    Dim subEnum As ElementEnumerator
    Set subEnum = cxEl.GetSubElements()
    Dim dTotal  As Double
    Dim oSub    As Element
    Do While subEnum.MoveNext
        Set oSub = subEnum.Current
        dTotal   = dTotal + PartialLengthPrimitive(oSub, zones)
    Loop
    PartialLengthComplexShape = dTotal
    Exit Function
ErrorHandler:
    PartialLengthComplexShape = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Length.PartialLengthComplexShape"
End Function