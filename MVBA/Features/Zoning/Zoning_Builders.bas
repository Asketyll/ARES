' Module: Zoning_Builders
' Description: Offset-buffer builders for the Zoning split - one shape per element type, plus the
'              rule that decides which caps are round. Pure geometry: no model access, no
'              module state, nothing read from config.
' Rationale, thresholds and the measurements behind them: _bmad/docs/zoning-mechanics.md
' License: This project is licensed under the AGPL-3.0.
' Dependencies: Geometry, ARESConstants, ErrorHandler

Option Explicit


' Round end-caps are built a hair wider than the offset they close, as a RATIO of that offset.
Private Const CAP_OVERLAP_RATIO As Double = 0.0005

' ============================================================
'  ZONE BUILDERS
'
'  Pure geometry functions.
'  Each builder returns a closed orphan Element (NOT added to the model).
'  The caller (dispatcher) passes the result to AddOrWrite.
' ============================================================

' BuildCellZone
' ---------------------------------------------------------------------------
' Creates a rounded rectangle aligned with the cell's own rotation axis.
Public Function BuildCellZone(ByVal oEl As Element, ByVal Dist As Double) As Element
    On Error GoTo ErrorHandler

    Dim cellEl    As CellElement
    Dim oRange    As Range3d         ' axis-aligned world bbox of the cell
    Dim invRot    As Matrix3d        ' inverse of cell rotation (= transpose for pure rotation)
    Dim oOrigin   As Point3d         ' cell insertion point in world space
    Dim corners(0 To 3) As Point3d   ' 4 corners of the world bbox
    Dim worldPt   As Point3d         ' one corner translated to origin (before inverse rotation)
    Dim localPt   As Point3d         ' same corner after inverse rotation (cell-local space)
    Dim xMin      As Double          ' local-space bbox extents
    Dim xMax      As Double
    Dim yMin      As Double
    Dim yMax      As Double
    Dim i         As Long
    Dim comps(0 To 7) As ChainableElement  ' 4 straight sides + 4 quarter-circle corner arcs
    Dim cxShape   As ComplexShapeElement
    Dim fwdT      As Transform3d     ' forward transform: local → world

    Set cellEl = oEl
    oOrigin    = cellEl.Origin                    ' world-space insertion point
    invRot     = Matrix3dInverse(cellEl.Rotation) ' inverse rotation matrix
    oRange     = oEl.Range                        ' world-space axis-aligned bbox

    ' Step 1: collect the 4 world-space corners of the bbox.
    corners(0) = Point3dFromXY(oRange.Low.X,  oRange.Low.Y)   ' bottom-left
    corners(1) = Point3dFromXY(oRange.High.X, oRange.Low.Y)   ' bottom-right
    corners(2) = Point3dFromXY(oRange.High.X, oRange.High.Y)  ' top-right
    corners(3) = Point3dFromXY(oRange.Low.X,  oRange.High.Y)  ' top-left

    ' Step 2 & 3: transform each corner to cell-local space and track extents.
    xMin = 1E+30 : xMax = -1E+30 : yMin = 1E+30 : yMax = -1E+30
    For i = 0 To 3
        ' Translate to origin so rotation is around (0,0), then apply inverse rotation.
        worldPt.X = corners(i).X - oOrigin.X
        worldPt.Y = corners(i).Y - oOrigin.Y
        worldPt.Z = corners(i).Z - oOrigin.Z
        localPt = Point3dFromMatrix3dTimesPoint3d(invRot, worldPt)
        If localPt.X < xMin Then xMin = localPt.X
        If localPt.X > xMax Then xMax = localPt.X
        If localPt.Y < yMin Then yMin = localPt.Y
        If localPt.Y > yMax Then yMax = localPt.Y
    Next i

    ' Step 4: build the 8-component rounded rectangle in cell-local space.
    ' Arc angle convention for CreateArcElement2: startAngle is measured CCW from
    ' the local X axis; sweepAngle is the signed arc span (positive = CCW).
    ' All 4 corner arcs sweep +PI/2 (quarter circle CCW).

    ' Bottom side: connects BL arc end  to BR arc start  (runs left → right at y=yMin-Dist)
    Set comps(0) = CreateLineElement2(Nothing, Point3dFromXY(xMin,       yMin - Dist), Point3dFromXY(xMax,       yMin - Dist))
    ' BR corner arc: center (xMax, yMin), starts pointing down (-PI/2) sweeps to pointing right
    Set comps(1) = CreateArcElement2(Nothing,  Point3dFromXY(xMax,       yMin),        Dist, Dist, Matrix3dIdentity, -Application.PI / 2, Application.PI / 2)
    ' Right side: connects BR arc end   to TR arc start  (runs bottom → top at x=xMax+Dist)
    Set comps(2) = CreateLineElement2(Nothing, Point3dFromXY(xMax + Dist, yMin),       Point3dFromXY(xMax + Dist, yMax))
    ' TR corner arc: center (xMax, yMax), starts pointing right (0) sweeps to pointing up
    Set comps(3) = CreateArcElement2(Nothing,  Point3dFromXY(xMax,       yMax),        Dist, Dist, Matrix3dIdentity,  0,                  Application.PI / 2)
    ' Top side: connects TR arc end     to TL arc start  (runs right → left at y=yMax+Dist)
    Set comps(4) = CreateLineElement2(Nothing, Point3dFromXY(xMax,       yMax + Dist), Point3dFromXY(xMin,       yMax + Dist))
    ' TL corner arc: center (xMin, yMax), starts pointing up (PI/2) sweeps to pointing left
    Set comps(5) = CreateArcElement2(Nothing,  Point3dFromXY(xMin,       yMax),        Dist, Dist, Matrix3dIdentity,  Application.PI / 2, Application.PI / 2)
    ' Left side: connects TL arc end    to BL arc start  (runs top → bottom at x=xMin-Dist)
    Set comps(6) = CreateLineElement2(Nothing, Point3dFromXY(xMin - Dist, yMax),       Point3dFromXY(xMin - Dist, yMin))
    ' BL corner arc: center (xMin, yMin), starts pointing left (PI) sweeps to pointing down
    Set comps(7) = CreateArcElement2(Nothing,  Point3dFromXY(xMin,       yMin),        Dist, Dist, Matrix3dIdentity,  Application.PI,     Application.PI / 2)

    ' CreateComplexShapeElement1 automatically reverses individual components as needed
    ' to ensure they form a single continuous closed loop.
    Set cxShape = CreateComplexShapeElement1(comps, msdFillModeNotFilled)

    ' Step 5: bring the shape back to world space.
    ' Transform3dFromMatrix3dPoint3d builds: P_world = Rotation * P_local + Origin
    fwdT = Transform3dFromMatrix3dPoint3d(cellEl.Rotation, oOrigin)
    cxShape.Transform fwdT

    Set BuildCellZone = cxShape
    Exit Function

ErrorHandler:
    Set BuildCellZone = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.BuildCellZone"
End Function

' BuildLineZone
' ---------------------------------------------------------------------------
' Creates a buffer zone around a straight line segment.
Public Function BuildLineZone(ByVal oEl As Element, _
                               ByVal Dist As Double, _
                               ByVal roundStart As Boolean, _
                               ByVal roundEnd As Boolean) As Element
    On Error GoTo ErrorHandler

    Dim lineEl As LineElement
    Dim ptS    As Point3d   ' segment start point
    Dim ptE    As Point3d   ' segment end point
    Dim perp   As Point3d   ' perpendicular offset vector (length = Dist, 90° CCW from S→E)
    Dim L0     As Point3d   ' left side start  (near ptS, offset left)
    Dim L1     As Point3d   ' left side end    (near ptE, offset left)
    Dim R0     As Point3d   ' right side start (near ptS, offset right)
    Dim R1     As Point3d   ' right side end   (near ptE, offset right)

    Set lineEl = oEl
    ptS  = lineEl.StartPoint
    ptE  = lineEl.EndPoint
    perp = Geometry.Perp2D(ptS, ptE, Dist)

    ' Guard: if the segment has zero length, Perp2D returns a zero vector.
    ' Point3dMagnitudeSquared returns |perp|^2; a valid perp has |perp|^2 = Dist^2 >> 1E-24.
    If Point3dMagnitudeSquared(perp) < 1E-24 Then Exit Function

    ' Compute the 4 rectangle corners using native MVBA Point3d arithmetic.
    L0 = Point3dAdd(ptS, perp)      : L1 = Point3dAdd(ptE, perp)       ' left side
    R1 = Point3dSubtract(ptE, perp) : R0 = Point3dSubtract(ptS, perp)  ' right side

    ' Fast path: both caps flat → a simple 4-corner rectangle (unchanged legacy behaviour).
    If Not roundStart And Not roundEnd Then
        Dim rectPts(0 To 4) As Point3d
        rectPts(0) = L0 : rectPts(1) = L1 : rectPts(2) = R1 : rectPts(3) = R0 : rectPts(4) = L0
        Set BuildLineZone = CreateShapeElement1(Nothing, rectPts)
        Exit Function
    End If

    ' Per-end caps → a ComplexShape running L0→L1→[end cap]→R1→R0→[start cap]→L0.
    '   end cap   (at ptE): semicircle L1→R1 (round) OR straight chord L1→R1 (flat)
    '   start cap (at ptS): semicircle R0→L0 (round) OR straight chord R0→L0 (flat)
    ' A round end-cap starts facing perp (toward L1) and sweeps -PI to R1; a round start-cap
    ' starts facing -perp (toward R0) and sweeps -PI to L0.
    Dim comps(0 To 3) As ChainableElement
    Set comps(0) = CreateLineElement2(Nothing, L0, L1)                                                  ' left side
    If roundEnd Then
        Set comps(1) = CreateArcElement2(Nothing, ptE, Dist * (1 + CAP_OVERLAP_RATIO), Dist * (1 + CAP_OVERLAP_RATIO), Matrix3dIdentity, Point3dPolarAngle(perp), -Application.PI)
    Else
        Set comps(1) = CreateLineElement2(Nothing, L1, R1)                                              ' flat end cap (chord)
    End If
    Set comps(2) = CreateLineElement2(Nothing, R1, R0)                                                  ' right side
    If roundStart Then
        Set comps(3) = CreateArcElement2(Nothing, ptS, Dist * (1 + CAP_OVERLAP_RATIO), Dist * (1 + CAP_OVERLAP_RATIO), Matrix3dIdentity, Point3dPolarAngle(Point3dNegate(perp)), -Application.PI)
    Else
        Set comps(3) = CreateLineElement2(Nothing, R0, L0)                                              ' flat start cap (chord)
    End If
    Set BuildLineZone = CreateComplexShapeElement1(comps, msdFillModeNotFilled)
    Exit Function

ErrorHandler:
    Set BuildLineZone = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.BuildLineZone"
End Function

' CapRoundAt
' ---------------------------------------------------------------------------
' Decides whether the buffer cap at a piece endpoint must be ROUNDED (smooth round-join) or
' FLAT. A cap is flat only at a FREE END of an open chain -- a point coincident with the chain's
' global Start or End. Every other endpoint is an interior junction and is rounded so flat-cap
' buffers are not cropped at sharp intermediate angles. When globalRoundOrClosed is True (the
' user asked for round caps everywhere, or the element is a closed shape with no free end) every
' cap is rounded and gStart/gEnd are ignored.
Public Function CapRoundAt(ByRef pt As Point3d, _
                            ByRef gStart As Point3d, _
                            ByRef gEnd As Point3d, _
                            ByVal globalRoundOrClosed As Boolean, _
                            ByVal tol As Double) As Boolean
    If globalRoundOrClosed Then
        CapRoundAt = True
    ElseIf Point3dEqualTolerance(pt, gStart, tol) Or Point3dEqualTolerance(pt, gEnd, tol) Then
        CapRoundAt = False    ' free chain end → flat cap
    Else
        CapRoundAt = True     ' interior junction → rounded cap
    End If
End Function

' BuildArcZone
' ---------------------------------------------------------------------------
' Creates a buffer zone around an arc element.
Public Function BuildArcZone(ByVal oEl As Element, _
                              ByVal Dist As Double, _
                              ByVal roundStart As Boolean, _
                              ByVal roundEnd As Boolean) As Element
    On Error GoTo ErrorHandler

    Dim arcEl           As ArcElement
    Dim outerArc        As ArcElement    ' source arc scaled outward by Dist
    Dim innerArc        As ArcElement    ' source arc scaled inward  by Dist (reversed)
    Dim capEnd          As ArcElement    ' full semicircle cap at arc end point   (only if roundEnd)
    Dim capStart        As ArcElement    ' full semicircle cap at arc start point (only if roundStart)
    Dim trimmedCapEnd   As ArcElement    ' cap trimmed to intersection (Case C)
    Dim trimmedCapStart As ArcElement
    Dim oCenter         As Point3d
    Dim rOuter          As Double        ' outer buffer radius = arcRadius + Dist
    Dim rInner          As Double        ' inner buffer radius = arcRadius - Dist (may be <= 0)
    Dim startAngle      As Double
    Dim sweepAngle      As Double
    Dim capSweep        As Double        ' sweep sign matches the original arc direction
    Dim ptOuterStart    As Point3d
    Dim ptOuterEnd      As Point3d
    Dim ptInnerStart    As Point3d
    Dim ptInnerEnd      As Point3d
    Dim ptArcStart      As Point3d       ' start point of the original arc
    Dim ptArcEnd        As Point3d       ' end   point of the original arc
    Dim isectPts()      As Point3d       ' intersection points between the two cap circles
    Dim nIsect          As Long          ' upper bound of isectPts (-1 if empty)
    Dim ptIsect         As Point3d       ' chosen intersection point (outermost)
    Dim dq0             As Double        ' squared distance from center to isectPts(0)
    Dim dq1             As Double        ' squared distance from center to isectPts(1)
    Dim angCES          As Double        ' capEnd   start angle
    Dim angCEE          As Double        ' capEnd   end   angle (at intersection)
    Dim angCSS          As Double        ' capStart start angle (at intersection)
    Dim angCSE          As Double        ' capStart end   angle
    Dim cxShape         As ComplexShapeElement
    Dim comps4(0 To 3)  As ChainableElement
    Dim parts()         As ChainableElement   ' pie case: variable-length ordered boundary
    Dim np              As Long
    Dim ptEndHub        As Point3d            ' pie case: end-side point at/near the center
    Dim ptStartHub      As Point3d            ' pie case: start-side point at/near the center

    Set arcEl    = oEl
    oCenter      = arcEl.CenterPoint
    rOuter       = arcEl.PrimaryRadius + Dist
    rInner       = arcEl.PrimaryRadius - Dist
    startAngle   = arcEl.StartAngle
    sweepAngle   = arcEl.SweepAngle

    ' Guard: zero-sweep arc cannot produce a valid zone.
    If Abs(sweepAngle) < 1E-10 Then Exit Function

    ' Build the outer arc: clone the original, then scale its radius outward.
    ' ScaleUniform(center, factor) scales all geometry uniformly around a point.
    Set outerArc = arcEl.Clone
    outerArc.ScaleUniform oCenter, rOuter / arcEl.PrimaryRadius
    ptOuterStart = outerArc.StartPoint
    ptOuterEnd   = outerArc.EndPoint
    ptArcStart   = arcEl.StartPoint
    ptArcEnd     = arcEl.EndPoint

    ' capSweep = ±PI: a semicircle sweeping in the same rotational direction
    ' as the original arc (positive for CCW, negative for CW).
    capSweep = Sgn(sweepAngle) * Application.PI

    ' Build only the round caps that are actually requested.
    '
    ' Built CAP_OVERLAP_RATIO wider than the offset, exactly as BuildLineZone does, and for the same
    ' reason: a cap of radius exactly Dist is tangent to the neighbouring buffer's flank at the shared
    ' junction, and GetRegionUnion misbehaves on boundaries that touch without crossing. The line
    ' builder was given that overlap when it fixed the incomplete IC/OL zoning; the arc builder was
    ' not, and that asymmetry is measurable - a repair run fed the union two rebuilt ARC buffers of
    ' 114.38 m2 and got back 1270.64 m2 in and 1270.64 m2 out, the arcs dropped without a word.
    ' capEnd   begins facing outward (toward ptOuterEnd) and sweeps a half circle toward the inner edge.
    ' capStart begins facing inward  (toward oCenter)    and sweeps a half circle back to the outer edge.
    If roundEnd Then
        Set capEnd = CreateArcElement2(Nothing, ptArcEnd, Dist * (1 + CAP_OVERLAP_RATIO), Dist * (1 + CAP_OVERLAP_RATIO), Matrix3dIdentity, _
                                        Point3dPolarAngle(Point3dSubtract(ptOuterEnd, ptArcEnd)), capSweep)
    End If
    If roundStart Then
        Set capStart = CreateArcElement2(Nothing, ptArcStart, Dist * (1 + CAP_OVERLAP_RATIO), Dist * (1 + CAP_OVERLAP_RATIO), Matrix3dIdentity, _
                                          Point3dPolarAngle(Point3dSubtract(oCenter, ptArcStart)), capSweep)
    End If

    ' --- Case C: both caps round and overlapping (arc near 360°) ---
    ' GetIntersectionPoints returns an empty array (raises error on UBound) if no intersection.
    If roundStart And roundEnd Then
        isectPts = capEnd.GetIntersectionPoints(capStart, Matrix3dIdentity)
        nIsect = -1 : On Error Resume Next : nIsect = UBound(isectPts) : On Error GoTo 0

        If nIsect >= 0 Then
            ' The caps overlap → use a 3-component shape with trimmed caps.
            ' Two circles can intersect at up to 2 points; we want the one
            ' that is farthest from the arc center (the "outer" intersection).
            If nIsect >= 1 Then
                dq0 = (isectPts(0).X - oCenter.X) ^ 2 + (isectPts(0).Y - oCenter.Y) ^ 2
                dq1 = (isectPts(1).X - oCenter.X) ^ 2 + (isectPts(1).Y - oCenter.Y) ^ 2
                If dq0 >= dq1 Then ptIsect = isectPts(0) Else ptIsect = isectPts(1)
            Else
                ptIsect = isectPts(0)
            End If

            ' Compute the angle to the intersection point from each cap center,
            ' then normalise the sweep to the correct direction (same as capSweep sign).
            angCES = Point3dPolarAngle(Point3dSubtract(ptOuterEnd,   ptArcEnd))
            angCEE = Point3dPolarAngle(Point3dSubtract(ptIsect,      ptArcEnd))
            angCSS = Point3dPolarAngle(Point3dSubtract(ptIsect,      ptArcStart))
            angCSE = Point3dPolarAngle(Point3dSubtract(ptOuterStart, ptArcStart))

            Set trimmedCapEnd   = CreateArcElement2(Nothing, ptArcEnd,   Dist * (1 + CAP_OVERLAP_RATIO), Dist * (1 + CAP_OVERLAP_RATIO), Matrix3dIdentity, _
                                                     angCES, Geometry.NormalizeAngle(angCEE - angCES, capSweep))
            Set trimmedCapStart = CreateArcElement2(Nothing, ptArcStart, Dist * (1 + CAP_OVERLAP_RATIO), Dist * (1 + CAP_OVERLAP_RATIO), Matrix3dIdentity, _
                                                     angCSS, Geometry.NormalizeAngle(angCSE - angCSS, capSweep))

            Dim compsO(0 To 2) As ChainableElement
            Set compsO(0) = outerArc
            Set compsO(1) = trimmedCapEnd    ' outer arc end → intersection point
            Set compsO(2) = trimmedCapStart  ' intersection point → outer arc start
            Set BuildArcZone = CreateComplexShapeElement1(compsO, msdFillModeNotFilled)
            Exit Function
        End If
    End If

    ' --- Case A or B: no cap overlap ---
    If rInner > 0 Then
        ' Case A — Annular sector: inner radius is positive, zone is a ring slice.
        ' Boundary: outerArc → [end cap] → innerArc → [start cap].
        ' A round cap reuses capEnd/capStart; a flat cap is the radial line that the cap replaces
        ' (capEnd spans ptOuterEnd→ptInnerStart, capStart spans ptInnerEnd→ptOuterStart).
        Set innerArc = arcEl.Clone
        innerArc.ScaleUniform oCenter, rInner / arcEl.PrimaryRadius
        ' Reverse the inner arc so the boundary runs as a continuous closed loop:
        '   outerArc goes start→end; innerArc must go end→start.
        innerArc.StartAngle = startAngle + sweepAngle
        innerArc.SweepAngle = -sweepAngle
        ptInnerStart = innerArc.StartPoint
        ptInnerEnd   = innerArc.EndPoint

        Set comps4(0) = outerArc
        If roundEnd Then
            Set comps4(1) = capEnd
        Else
            Set comps4(1) = CreateLineElement2(Nothing, ptOuterEnd, ptInnerStart)
        End If
        Set comps4(2) = innerArc
        If roundStart Then
            Set comps4(3) = capStart
        Else
            Set comps4(3) = CreateLineElement2(Nothing, ptInnerEnd, ptOuterStart)
        End If
        Set cxShape = CreateComplexShapeElement1(comps4, msdFillModeNotFilled)
    Else
        ' Case B — Pie sector: Dist >= arc radius, inner arc collapses toward the center.
        ' Each side reaches the center either via its round cap (landing near, not at, the center)
        ' or via a straight radial line (landing exactly at the center). A short bridge line joins
        ' the two hub points, and is skipped when they coincide (both flat → both at oCenter).
        If roundEnd Then ptEndHub = capEnd.EndPoint Else ptEndHub = oCenter
        If roundStart Then ptStartHub = capStart.StartPoint Else ptStartHub = oCenter

        np = 0
        ReDim parts(0 To 3)
        Set parts(np) = outerArc : np = np + 1
        If roundEnd Then
            Set parts(np) = capEnd
        Else
            Set parts(np) = CreateLineElement2(Nothing, ptOuterEnd, oCenter)
        End If
        np = np + 1
        If Not Point3dEqualTolerance(ptEndHub, ptStartHub, Dist * 0.000000001) Then
            Set parts(np) = CreateLineElement2(Nothing, ptEndHub, ptStartHub) : np = np + 1
        End If
        If roundStart Then
            Set parts(np) = capStart
        Else
            Set parts(np) = CreateLineElement2(Nothing, oCenter, ptOuterStart)
        End If
        np = np + 1
        ReDim Preserve parts(0 To np - 1)
        Set cxShape = CreateComplexShapeElement1(parts, msdFillModeNotFilled)
    End If

    Set BuildArcZone = cxShape
    Exit Function

ErrorHandler:
    Set BuildArcZone = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.BuildArcZone"
End Function
