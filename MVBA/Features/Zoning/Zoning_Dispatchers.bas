' Module: Zoning_Dispatchers
' Description: One dispatcher per element type for the Zoning split. Each calls the matching
'              builder and hands the result to AddOrWrite, which decides whether it is
'              stored for the global merge or written straight out.
' Rationale, thresholds and the measurements behind them: _bmad/docs/zoning-mechanics.md
' License: This project is licensed under the AGPL-3.0.
' Dependencies: Zoning (AddOrWrite, FuseRegions, WriteDebugClones), Zoning_Builders, ErrorHandler

Option Explicit


' DispatchElement
' ---------------------------------------------------------------------------
' Routes ONE source element to the dispatcher for its type. Extracted so the coverage repair can
' rebuild an element's buffer exactly the way the first pass built it - the alternative was a second
' decomposition of lines, arcs and chains living beside the first and drifting away from it.
Public Sub DispatchElement(ByVal oEl As Element, _
                            ByVal Dist As Double, _
                            ByVal TargetLevel As Level, _
                            ByVal Color As Long, _
                            ByVal Style As String, _
                            ByVal Weight As Long, _
                            ByRef outBufs() As Element, _
                            ByRef nOut As Long, _
                            ByVal DebugMode As Boolean, _
                            ByVal RoundCaps As Boolean)
    On Error GoTo ErrorHandler
    Select Case oEl.Type
        Case msdElementTypeLine
            ZoneFromLine oEl, Dist, TargetLevel, Color, Style, Weight, outBufs, nOut, DebugMode, RoundCaps
        Case msdElementTypeLineString
            ZoneFromLineString oEl, Dist, TargetLevel, Color, Style, Weight, outBufs, nOut, DebugMode, RoundCaps
        Case msdElementTypeArc
            ZoneFromArc oEl, Dist, TargetLevel, Color, Style, Weight, outBufs, nOut, DebugMode, RoundCaps
        Case msdElementTypeComplexString, msdElementTypeComplexShape
            ZoneFromComplexString oEl, Dist, TargetLevel, Color, Style, Weight, outBufs, nOut, DebugMode, RoundCaps
        Case msdElementTypeEllipse
            ZoneFromEllipse oEl, Dist, TargetLevel, Color, Style, Weight, outBufs, nOut, DebugMode
        Case msdElementTypeCellHeader
            ZoneFromCell oEl, Dist, TargetLevel, Color, Style, Weight, outBufs, nOut, DebugMode
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.DispatchElement"
End Sub

' ============================================================
'  ZONE DISPATCHERS
'
'  One dispatcher per element type.
'  Responsibility: call the matching builder, then pass the result
'  to AddOrWrite (which decides whether to store or write it).
'
'  Signature pattern shared by all dispatchers:
'    oEl         → the source element
'    Dist        → buffer distance
'    TargetLevel / Color / Style / Weight → output symbology
'    outBufs / nOut → the accumulator array and its sentinel counter
' ============================================================

' ZoneFromLine
' Handles a single straight line segment (msdElementTypeLine).
' Produces one stadium shape (rectangle + semicircular end-caps).
Private Sub ZoneFromLine(ByVal oEl As Element, _
                         ByVal Dist As Double, _
                         ByVal TargetLevel As Level, _
                         ByVal Color As Long, _
                         ByVal Style As String, _
                         ByVal Weight As Long, _
                         ByRef outBufs() As Element, _
                         ByRef nOut As Long, _
                         ByVal DebugMode As Boolean, _
                         ByVal RoundCaps As Boolean)
    On Error GoTo ErrorHandler
    Dim elem As Element
    ' Single segment: both ends are free ends of the chain → caps follow the global RoundCaps flag.
    Set elem = BuildLineZone(oEl, Dist, RoundCaps, RoundCaps)
    If Not elem Is Nothing Then AddOrWrite elem, TargetLevel, Color, Style, Weight, outBufs, nOut, Dist
    Exit Sub
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.ZoneFromLine"
End Sub

' ZoneFromLineString
' Handles a polyline element (msdElementTypeLineString).
Private Sub ZoneFromLineString(ByVal oEl As Element, _
                               ByVal Dist As Double, _
                               ByVal TargetLevel As Level, _
                               ByVal Color As Long, _
                               ByVal Style As String, _
                               ByVal Weight As Long, _
                               ByRef outBufs() As Element, _
                               ByRef nOut As Long, _
                               ByVal DebugMode As Boolean, _
                               ByVal RoundCaps As Boolean)
    On Error GoTo ErrorHandler

    Dim oVL       As VertexList  ' exposes vertex list of any VertexList-compatible element
    Dim v()       As Point3d     ' array of all vertices in the polyline
    Dim n         As Long        ' total number of vertices
    Dim j         As Long
    Dim subBufs() As Element     ' stadiums for each individual segment
    Dim nBuf      As Long        ' number of valid stadiums built so far
    Dim buf       As Element
    Dim gStart    As Point3d     ' polyline global start (free end candidate)
    Dim gEnd      As Point3d     ' polyline global end   (free end candidate)
    Dim allRound  As Boolean     ' True → every cap rounded (global RoundCaps, or closed polyline)
    Dim tol       As Double

    Set oVL = oEl
    v = oVL.GetVertices
    n = UBound(v) - LBound(v) + 1
    If n < 2 Then Exit Sub   ' nothing to buffer with fewer than 2 vertices

    ' Caps are flat only at the polyline's two global ends (v(0), v(n-1)); every interior vertex
    ' gets a rounded round-join so flat-cap buffers are not cropped at sharp angles. A closed
    ' polyline (v(0) == v(n-1)) has no free end → every cap rounded.
    tol      = Dist * ARES_CAP_MATCH_FRAC
    gStart   = v(0)
    gEnd     = v(n - 1)
    allRound = RoundCaps Or Point3dEqualTolerance(gStart, gEnd, tol)

    ' Step 1: build one stadium per segment, choosing each cap by free-end test.
    nBuf = 0
    For j = 0 To n - 2
        ' CreateLineElement2(Nothing, ...) creates a temporary line not added to the model.
        Set buf = BuildLineZone(CreateLineElement2(Nothing, v(j), v(j + 1)), Dist, _
                                CapRoundAt(v(j),     gStart, gEnd, allRound, tol), _
                                CapRoundAt(v(j + 1), gStart, gEnd, allRound, tol))
        If Not buf Is Nothing Then
            ReDim Preserve subBufs(0 To nBuf)
            Set subBufs(nBuf) = buf
            nBuf = nBuf + 1
        End If
    Next j

    If nBuf = 0 Then Exit Sub

    If DebugMode Then WriteDebugClones subBufs, nBuf, TargetLevel, Color, Style, Weight

    ' Step 2: fuse the per-segment stadiums into clean region(s) and emit.
    Dim merged() As Element
    Dim nMerged  As Long
    FuseRegions subBufs, nBuf, merged, nMerged, DebugMode, "linestring id=" & DLongToString(oEl.ID)
    For j = 0 To nMerged - 1
        AddOrWrite merged(j), TargetLevel, Color, Style, Weight, outBufs, nOut, Dist
    Next j
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.ZoneFromLineString"
End Sub

' ZoneFromArc
' Handles a single arc element (msdElementTypeArc).
' Produces an annular sector (ring slice) when Dist < arc radius,
' or a pie sector when Dist >= arc radius.
Private Sub ZoneFromArc(ByVal oEl As Element, _
                        ByVal Dist As Double, _
                        ByVal TargetLevel As Level, _
                        ByVal Color As Long, _
                        ByVal Style As String, _
                        ByVal Weight As Long, _
                        ByRef outBufs() As Element, _
                        ByRef nOut As Long, _
                        ByVal DebugMode As Boolean, _
                        ByVal RoundCaps As Boolean)
    On Error GoTo ErrorHandler
    Dim elem As Element
    ' Single arc: both ends are free ends of the chain → caps follow the global RoundCaps flag.
    Set elem = BuildArcZone(oEl, Dist, RoundCaps, RoundCaps)
    If Not elem Is Nothing Then AddOrWrite elem, TargetLevel, Color, Style, Weight, outBufs, nOut, Dist
    Exit Sub
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.ZoneFromArc"
End Sub

' ZoneFromComplexString
' Handles ComplexString and ComplexShape elements.
' These are chains of sub-elements (lines, arcs, and nested linestrings).
Private Sub ZoneFromComplexString(ByVal oEl As Element, _
                                  ByVal Dist As Double, _
                                  ByVal TargetLevel As Level, _
                                  ByVal Color As Long, _
                                  ByVal Style As String, _
                                  ByVal Weight As Long, _
                                  ByRef outBufs() As Element, _
                                  ByRef nOut As Long, _
                                  ByVal DebugMode As Boolean, _
                                  ByVal RoundCaps As Boolean)
    On Error GoTo ErrorHandler

    ' ComplexElement is the common interface for both ComplexStringElement and
    ' ComplexShapeElement. Using it here lets us handle closed loops (stored as
    ' ComplexShapeElement) without Error 91 on the implicit interface cast.
    Dim cxEl      As ComplexElement
    Dim subEnum   As ElementEnumerator
    Dim comp      As Element    ' current sub-element being processed
    Dim buf       As Element    ' stadium / sector for comp
    Dim subBufs() As Element    ' all stadiums/sectors accumulated before fusion
    Dim nBuf      As Long
    Dim oVLs      As VertexList ' used to read vertices of a LineString sub-element
    Dim vs()      As Point3d
    Dim ns        As Long
    Dim js        As Long
    Dim j         As Long
    Dim gStart    As Point3d    ' chain global start (free end candidate; unused when closed)
    Dim gEnd      As Point3d    ' chain global end   (free end candidate; unused when closed)
    Dim allRound  As Boolean    ' True → every cap rounded (global RoundCaps, or closed shape)
    Dim tol       As Double

    Set cxEl    = oEl
    Set subEnum = cxEl.GetSubElements()
    nBuf = 0
    tol  = Dist * ARES_CAP_MATCH_FRAC

    ' Free-end detection. A ComplexShape is always closed → no free end, every cap rounded.
    ' A ComplexString is an open chain whose global Start/End points are its two free ends.
    ' (A degenerate ComplexString with Start == End is treated as closed → every cap rounded.)
    If oEl.Type = msdElementTypeComplexShape Then
        allRound = True
    Else
        gStart   = oEl.AsChainableElement.StartPoint
        gEnd     = oEl.AsChainableElement.EndPoint
        allRound = RoundCaps Or Point3dEqualTolerance(gStart, gEnd, tol)
    End If

    Do While subEnum.MoveNext
        Set comp = subEnum.Current
        Set buf  = Nothing   ' reset for each sub-element

        Select Case comp.Type
            Case msdElementTypeLine
                Set buf = BuildLineZone(comp, Dist, _
                            CapRoundAt(comp.AsChainableElement.StartPoint, gStart, gEnd, allRound, tol), _
                            CapRoundAt(comp.AsChainableElement.EndPoint,   gStart, gEnd, allRound, tol))

            Case msdElementTypeLineString
                ' Expand into per-segment stadiums to handle self-crossing polylines
                ' (same strategy as ZoneFromLineString). Interior LineString vertices never match
                ' the chain ends, so they are always rounded; only a vertex coincident with the
                ' chain's global Start/End (a free end) gets a flat cap.
                Set oVLs = comp
                vs = oVLs.GetVertices
                ns = UBound(vs) - LBound(vs) + 1
                For js = 0 To ns - 2
                    Set buf = BuildLineZone(CreateLineElement2(Nothing, vs(js), vs(js + 1)), Dist, _
                                CapRoundAt(vs(js),     gStart, gEnd, allRound, tol), _
                                CapRoundAt(vs(js + 1), gStart, gEnd, allRound, tol))
                    If Not buf Is Nothing Then
                        ReDim Preserve subBufs(0 To nBuf)
                        Set subBufs(nBuf) = buf
                        nBuf = nBuf + 1
                    End If
                Next js
                Set buf = Nothing   ' already added above → skip the generic add below

            Case msdElementTypeArc
                Set buf = BuildArcZone(comp, Dist, _
                            CapRoundAt(comp.AsChainableElement.StartPoint, gStart, gEnd, allRound, tol), _
                            CapRoundAt(comp.AsChainableElement.EndPoint,   gStart, gEnd, allRound, tol))
        End Select

        ' Generic add for Line and Arc cases (buf is Nothing for LineString).
        If Not buf Is Nothing Then
            ReDim Preserve subBufs(0 To nBuf)
            Set subBufs(nBuf) = buf
            nBuf = nBuf + 1
        End If
    Loop

    If nBuf = 0 Then Exit Sub

    If DebugMode Then WriteDebugClones subBufs, nBuf, TargetLevel, Color, Style, Weight

    ' Fuse all sub-element buffers into clean region(s) and emit.
    Dim merged() As Element
    Dim nMerged  As Long
    FuseRegions subBufs, nBuf, merged, nMerged, DebugMode, "complexstring id=" & DLongToString(oEl.ID)
    For j = 0 To nMerged - 1
        AddOrWrite merged(j), TargetLevel, Color, Style, Weight, outBufs, nOut, Dist
    Next j
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.ZoneFromComplexString"
End Sub

' ZoneFromEllipse
' Handles EllipseElement (circles and ellipses — MicroStation stores both as EllipseElement).
Private Sub ZoneFromEllipse(ByVal oEl As Element, _
                            ByVal Dist As Double, _
                            ByVal TargetLevel As Level, _
                            ByVal Color As Long, _
                            ByVal Style As String, _
                            ByVal Weight As Long, _
                            ByRef outBufs() As Element, _
                            ByRef nOut As Long, _
                            ByVal DebugMode As Boolean)
    On Error GoTo ErrorHandler

    Dim ellEl         As EllipseElement
    Dim outerEl       As EllipseElement
    Dim innerEl       As EllipseElement
    Dim solid(0 To 0) As Element
    Dim holes(0 To 0) As Element
    Dim oEnum         As ElementEnumerator

    Set ellEl = oEl

    ' Build the outer ellipse: expand both radii by Dist, preserve center and rotation.
    Set outerEl = CreateEllipseElement2(Nothing, _
                                         ellEl.CenterPoint, _
                                         ellEl.PrimaryRadius   + Dist, _
                                         ellEl.SecondaryRadius + Dist, _
                                         ellEl.Rotation, _
                                         msdFillModeNotFilled)

    If (ellEl.PrimaryRadius - Dist) > 0 And (ellEl.SecondaryRadius - Dist) > 0 Then
        ' Case A: subtract the inner ellipse → annular (donut) planar region.
        Set innerEl = CreateEllipseElement2(Nothing, _
                                             ellEl.CenterPoint, _
                                             ellEl.PrimaryRadius   - Dist, _
                                             ellEl.SecondaryRadius - Dist, _
                                             ellEl.Rotation, _
                                             msdFillModeNotFilled)
        Set solid(0) = outerEl
        Set holes(0) = innerEl
        Set oEnum = GetRegionDifference(solid, holes, Nothing, msdFillModeNotFilled)
        If Not oEnum Is Nothing Then
            Do While oEnum.MoveNext
                AddOrWrite oEnum.Current, TargetLevel, Color, Style, Weight, outBufs, nOut, Dist
            Loop
        End If
    Else
        ' Case B: inner ellipse would have zero or negative radius → outer ellipse only.
        AddOrWrite outerEl, TargetLevel, Color, Style, Weight, outBufs, nOut, Dist
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.ZoneFromEllipse"
End Sub

' ZoneFromCell
' Handles CellHeader elements (placed blocks / symbols).
' Builds a rotated rounded rectangle around the cell's bounding box,
' aligned with the cell's own rotation (not world-axis-aligned).
Private Sub ZoneFromCell(ByVal oEl As Element, _
                         ByVal Dist As Double, _
                         ByVal TargetLevel As Level, _
                         ByVal Color As Long, _
                         ByVal Style As String, _
                         ByVal Weight As Long, _
                         ByRef outBufs() As Element, _
                         ByRef nOut As Long, _
                         ByVal DebugMode As Boolean)
    On Error GoTo ErrorHandler
    Dim elem As Element
    Set elem = BuildCellZone(oEl, Dist)
    If Not elem Is Nothing Then AddOrWrite elem, TargetLevel, Color, Style, Weight, outBufs, nOut, Dist
    Exit Sub
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.ZoneFromCell"
End Sub
