' Zoning.bas
' Description: Generates a "buffer zone" (safety boundary / offset shape) around elements
' found on specified levels.
'
' SUPPORTED ELEMENT TYPES
'   Line                        → stadium shape (rectangle + two semicircular end-caps)
'   LineString                  → one stadium per segment, all fused via GetRegionUnion
'   Arc                         → annular sector or pie sector (rounded or flat caps)
'   ComplexString / ComplexShape → same fusion strategy, one buffer per sub-element
'   CellHeader                  → rotated rounded rectangle aligned with the cell's own axis
'   EllipseElement (circle/ellipse)
'
' HOW IT WORKS
'   1. Collect all matching elements from the active model.
'   2. Dispatch each element to its typed zone builder.
'      Each builder returns an orphan closed shape — it is NOT added to the model.
'   3. Accumulate all zones, fuse them into a single region with GetRegionUnion, then write the result.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, ARESConstants, ErrorHandlerClass, Geometry, GetElements

Option Explicit

' Generates offset zones around elements on the specified source levels.
'
' Parameters (all optional — ARESConfig values are used when omitted):
'   Lvls        : source level name(s).
'                 Accepts: a single String, a String array, or omitted/empty
'                 (falls back to ARES_ZONING_LEVEL config value).
'   OutputLevel : name of the level that receives the new zone elements.
'   Color       : color index for the zone elements  (-1 = use config default).
'   Style       : line-style name for the zone elements ("" = use config default).
'   Weight      : line weight for the zone elements   (-1 = use config default).
'   Dist        : buffer distance in master units      (0  = use config default).
'   MergeZones  : True  (default) → fuse all individual zones together with
'                                   GetRegionUnion before writing to the model.
'                 False           → write each element's zone separately.
'   DebugMode   : True → write each individual zone shape to the model before
'                 the final merge, making pre-merge buffers visible alongside
'                 the merged result. Intended for geometry debugging. Default False.
'   RoundCaps   : True  (default) → open buffers (line / arc / linestring /
'                                   complexstring) get semicircular end-caps.
'                 False           → open buffers get flat (square / radial) caps.
'                 Closed elements (cell, ellipse) have no open cap and ignore this.
Public Sub Zoning(Optional Lvls As Variant, _
                  Optional OutputLevel As String = "", _
                  Optional Color As Long = -1, _
                  Optional Style As String = "", _
                  Optional Weight As Long = -1, _
                  Optional Dist As Double = 0, _
                  Optional MergeZones As Boolean = True, _
                  Optional DebugMode As Boolean = False, _
                  Optional RoundCaps As Boolean = True)

    On Error GoTo ErrorHandler

    Dim TargetLevel As Level
    Dim Elements()  As Element
    Dim i           As Long
    Dim k           As Long
    Dim oEl         As Element
    Dim allBufs()   As Element  ' accumulator used when MergeZones = True
    Dim nAllBufs    As Long     ' sentinel: -1 = write immediately; >=0 = accumulate

    ' --- Guard: configuration must be initialised before we can read config vars ---
    If Not ARESConfig.IsInitialized Then
        ErrorHandler.HandleError "ARESConfig not initialized", 0, "", "Zoning.Zoning"
        Exit Sub
    End If

    ' --- Fill in any missing parameters from ARESConfig ---
    If Len(OutputLevel) = 0 Then OutputLevel = ARESConfig.ARES_ZONING_OUTPUT_LEVEL.Value
    If Color  = -1          Then Color        = CLng(ARESConfig.ARES_ZONING_OUTPUT_COLOR.Value)
    If Len(Style) = 0       Then Style        = ARESConfig.ARES_ZONING_OUTPUT_STYLE.Value
    If Weight = -1          Then Weight       = CLng(ARESConfig.ARES_ZONING_OUTPUT_WEIGHT.Value)
    If Dist   <= 0          Then Dist         = Val(ARESConfig.ARES_ZONING_DISTANCE.Value)

    ' --- Resolve the source level list into a String array ---
    ' We accept three forms: omitted/empty → read from config;
    '                        a single String → wrap in a 1-element array;
    '                        a String array  → copy as-is.
    Dim ResolvedLvls() As String
    ' IMPORTANT: test IsArray FIRST. VBA does not short-circuit Or/And, so any CStr(Lvls) in the
    ' later branches is evaluated even when Lvls is an array — and CStr(array) raises Error 13
    ' (type mismatch). RunOutline passes a String() array here; RunZoning omits Lvls (Missing) and
    ' falls through to the config branch, where CStr on a Missing/scalar is safe.
    If IsArray(Lvls) Then
        ReDim ResolvedLvls(LBound(Lvls) To UBound(Lvls))
        For k = LBound(Lvls) To UBound(Lvls)
            ResolvedLvls(k) = CStr(Lvls(k))
        Next k
    ElseIf IsMissing(Lvls) Or IsEmpty(Lvls) Or Len(Trim(CStr(Lvls))) = 0 Then
        Dim LvlsStr As String
        LvlsStr = ARESConfig.ARES_ZONING_LEVEL.Value
        If Len(LvlsStr) = 0 Then
            ErrorHandler.HandleError "No levels provided and ARES_Zoning_Level config is empty", 0, "", "Zoning.Zoning"
            Exit Sub
        End If
        ResolvedLvls = Split(LvlsStr, ARES_VAR_DELIMITER)
    Else
        ReDim ResolvedLvls(0 To 0)
        ResolvedLvls(0) = CStr(Lvls)
    End If

    ' --- Validate the final parameter values ---
    If Dist <= 0 Then
        ErrorHandler.HandleError "Distance must be greater than zero", 0, "", "Zoning.Zoning"
        Exit Sub
    End If
    If UBound(ResolvedLvls) < LBound(ResolvedLvls) Then
        ErrorHandler.HandleError "No levels provided", 0, "", "Zoning.Zoning"
        Exit Sub
    End If
    If Not Application.HasActiveModelReference Then
        ErrorHandler.HandleError "No active model reference", 0, "", "Zoning.Zoning"
        Exit Sub
    End If

    ' --- Get (or create) the output level ---
    Set TargetLevel = GetElements.GetLevel(OutputLevel)
    If TargetLevel Is Nothing Then
        ErrorHandler.HandleError "Failed to get or create output level: " & OutputLevel, 0, "", "Zoning.Zoning"
        Exit Sub
    End If

    ' --- Collect all source elements by level and type ---
    Dim ee As ElementEnumerator
    Set ee = GetElements.ByEE(Levels:=ResolvedLvls, _
                              ElTypes:=Array(msdElementTypeCellHeader, _
                                            msdElementTypeLine, _
                                            msdElementTypeLineString, _
                                            msdElementTypeShape, _
                                            msdElementTypeComplexString, _
                                            msdElementTypeComplexShape, _
                                            msdElementTypeArc, _
                                            msdElementTypeEllipse))
    Elements = ee.BuildArrayFromContents

    If IsArray(Elements) Then
        If UBound(Elements) < LBound(Elements) Then
            ErrorHandler.HandleError "No elements found on specified levels", 0, "", "Zoning.Zoning"
            Exit Sub
        End If
    Else
        ErrorHandler.HandleError "Failed to retrieve elements", 0, "", "Zoning.Zoning"
        Exit Sub
    End If

    ' --- Set the output strategy via the nAllBufs sentinel ---
    ' nAllBufs = -1  → AddOrWrite will call WriteEl immediately (MergeZones = False)
    ' nAllBufs >= 0  → AddOrWrite accumulates into allBufs(); merge happens below
    If MergeZones Then nAllBufs = 0 Else nAllBufs = -1

    ' --- Process each element ---
    For i = LBound(Elements) To UBound(Elements)
        Set oEl = Elements(i)
        Select Case oEl.Type
            Case msdElementTypeLine
                ZoneFromLine oEl, Dist, TargetLevel, Color, Style, Weight, allBufs, nAllBufs, DebugMode, RoundCaps
            Case msdElementTypeLineString
                ZoneFromLineString oEl, Dist, TargetLevel, Color, Style, Weight, allBufs, nAllBufs, DebugMode, RoundCaps
            Case msdElementTypeArc
                ZoneFromArc oEl, Dist, TargetLevel, Color, Style, Weight, allBufs, nAllBufs, DebugMode, RoundCaps
            Case msdElementTypeComplexString, msdElementTypeComplexShape
                ZoneFromComplexString oEl, Dist, TargetLevel, Color, Style, Weight, allBufs, nAllBufs, DebugMode, RoundCaps
            Case msdElementTypeEllipse
                ZoneFromEllipse oEl, Dist, TargetLevel, Color, Style, Weight, allBufs, nAllBufs, DebugMode
            Case msdElementTypeCellHeader
                ZoneFromCell oEl, Dist, TargetLevel, Color, Style, Weight, allBufs, nAllBufs, DebugMode
        End Select
    Next i

    ' --- Merge all accumulated zones and write to the model (MergeZones = True only) ---
    If MergeZones And nAllBufs > 0 Then
        ' Debug mode: write a clone of each pre-merge shape so the individual zones
        ' are visible alongside the final merged result.
        If DebugMode Then WriteDebugClones allBufs, nAllBufs, TargetLevel, Color, Style, Weight

        Dim mergedAll() As Element
        Dim nMergedAll  As Long
        FuseRegions allBufs, nAllBufs, mergedAll, nMergedAll
        For k = 0 To nMergedAll - 1
            WriteEl mergedAll(k), TargetLevel, Color, Style, Weight
        Next k
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.Zoning"
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
    If Not elem Is Nothing Then AddOrWrite elem, TargetLevel, Color, Style, Weight, outBufs, nOut
    Exit Sub
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.ZoneFromLine"
End Sub

' ZoneFromLineString
' Handles a polyline element (msdElementTypeLineString).
'
' WHY NOT BUILD ONE SHAPE FOR THE WHOLE POLYLINE?
' A single offset of a self-crossing polyline (figure-4, figure-8) produces
' a self-intersecting outline. MicroStation's GetRegionUnion cannot fuse a
' self-intersecting shape into a clean region.
'
' STRATEGY: treat each segment independently.
'   1. Build a stadium (round-cap rectangle) for every segment.
'   2. Fuse all stadiums with GetRegionUnion.
' Because each stadium is a valid convex shape, GetRegionUnion always
' produces a clean non-self-intersecting result.
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
    FuseRegions subBufs, nBuf, merged, nMerged
    For j = 0 To nMerged - 1
        AddOrWrite merged(j), TargetLevel, Color, Style, Weight, outBufs, nOut
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
    If Not elem Is Nothing Then AddOrWrite elem, TargetLevel, Color, Style, Weight, outBufs, nOut
    Exit Sub
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.ZoneFromArc"
End Sub

' ZoneFromComplexString
' Handles ComplexString and ComplexShape elements.
' These are chains of sub-elements (lines, arcs, and nested linestrings).
'
' STRATEGY: same per-segment fusion used by ZoneFromLineString.
'   1. Iterate sub-elements via GetSubElements().
'   2. For each sub-element:
'      - Line   → one stadium.
'      - Arc    → one sector.
'      - LineString → expand further into per-segment stadiums
'                     (same self-crossing protection as ZoneFromLineString).
'   3. Fuse all results with GetRegionUnion.
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
    FuseRegions subBufs, nBuf, merged, nMerged
    For j = 0 To nMerged - 1
        AddOrWrite merged(j), TargetLevel, Color, Style, Weight, outBufs, nOut
    Next j
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.ZoneFromComplexString"
End Sub

' ZoneFromEllipse
' Handles EllipseElement (circles and ellipses — MicroStation stores both as EllipseElement).
'
' CASE A — Annular zone (both inner radii > 0):
'   outer = (PrimaryRadius + Dist, SecondaryRadius + Dist)
'   inner = (PrimaryRadius - Dist, SecondaryRadius - Dist)
'   GetRegionDifference(outer, inner) → donut-shaped planar ComplexShapeElement.
'
' CASE B — Full zone (at least one inner radius <= 0):
'   GetRegionDifference with an empty holes array returns a plain EllipseElement,
'   not a ComplexShapeElement — no benefit over writing outerEl directly.
'   The outer EllipseElement is already a closed planar element; written as-is.
'
' Approximation note: the exact offset curve of an ellipse is NOT an ellipse.
' Expanding both radii by Dist gives a uniform offset only for a circle; for a
' true ellipse the actual perimeter distance varies slightly. Acceptable for zoning.
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
                AddOrWrite oEnum.Current, TargetLevel, Color, Style, Weight, outBufs, nOut
            Loop
        End If
    Else
        ' Case B: inner ellipse would have zero or negative radius → outer ellipse only.
        AddOrWrite outerEl, TargetLevel, Color, Style, Weight, outBufs, nOut
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
    If Not elem Is Nothing Then AddOrWrite elem, TargetLevel, Color, Style, Weight, outBufs, nOut
    Exit Sub
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.ZoneFromCell"
End Sub

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
'
' CONSTRUCTION PIPELINE (all coordinates are in cell-local space until Step 5):
'
'   Step 1: Read the cell's world-space axis-aligned bounding box (Range3d).
'           A Range3d is always axis-aligned in world space, so a rotated cell
'           will have a larger bbox than its actual footprint.
'
'   Step 2: Transform the 4 bbox corners from world space to cell-local space
'           by translating to the cell origin, then multiplying by the inverse
'           rotation matrix.  This removes the cell's rotation so we work in a
'           simple axis-aligned coordinate system.
'
'   Step 3: Find the local-space extents (xMin, xMax, yMin, yMax).
'           Because a rotated bbox is larger than the actual footprint, we MUST
'           project all 4 corners — any one of them could be the min or max.
'
'   Step 4: Build the 8-component rounded rectangle in local space.
'           Arc centers sit at the 4 corners of the local bbox.
'           Each arc has radius = Dist and sweeps PI/2 (quarter circle).
'           The 4 straight sides connect adjacent arc endpoints.
'
'           Diagram (D = Dist, corners = arc centers):
'
'           (x0, y1+D) ────────── (x1, y1+D)
'          /                                  \
'   (x0-D, y1)  [TL arc]       [TR arc]  (x1+D, y1)
'       |                                       |
'   (x0-D, y0)  [BL arc]       [BR arc]  (x1+D, y0)
'          \                                  /
'           (x0, y0-D) ────────── (x1, y0-D)
'
'           Where x0=xMin, y0=yMin, x1=xMax, y1=yMax (local space).
'
'   Step 5: Apply the forward transform (Rotation * P_local + Origin) to bring
'           the shape back into world space with the correct rotation and position.
' ---------------------------------------------------------------------------
Private Function BuildCellZone(ByVal oEl As Element, ByVal Dist As Double) As Element
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
'
' FLAT caps  : 4-point closed rectangle (ShapeElement).
' ROUND caps : stadium shape — a ComplexShapeElement with:
'                - 2 straight sides parallel to the segment (offset by Dist left/right)
'                - 2 semicircular end-caps (radius = Dist), one at each endpoint.
'
'   Top view (round caps):
'          L0                L1
'          ╭────────────────╮
'         ╰                  ╯
'          ╰────────────────╯
'          R0       S     E  R1
'
'   Where S = segment start, E = segment end,
'   L = left side (offset by perp), R = right side (offset by -perp).
' ---------------------------------------------------------------------------
Private Function BuildLineZone(ByVal oEl As Element, _
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
        Set comps(1) = CreateArcElement2(Nothing, ptE, Dist, Dist, Matrix3dIdentity, Point3dPolarAngle(perp), -Application.PI)
    Else
        Set comps(1) = CreateLineElement2(Nothing, L1, R1)                                              ' flat end cap (chord)
    End If
    Set comps(2) = CreateLineElement2(Nothing, R1, R0)                                                  ' right side
    If roundStart Then
        Set comps(3) = CreateArcElement2(Nothing, ptS, Dist, Dist, Matrix3dIdentity, Point3dPolarAngle(Point3dNegate(perp)), -Application.PI)
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
' ---------------------------------------------------------------------------
Private Function CapRoundAt(ByRef pt As Point3d, _
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
'
' The outer and inner buffer arcs are built by cloning the source arc and
' uniformly scaling its radius around the arc center.
'
' CASE A — Annular sector (arc radius > Dist):
'   Both outer and inner arcs exist.
'   Shape = outerArc | cap_at_end | innerArc_reversed | cap_at_start
'
'   Top view (round caps):
'       ╭─── outerArc ───╮
'      ╰  cap           cap  ╯
'       ╰─── innerArc ───╯
'
' CASE B — Pie sector (arc radius <= Dist):
'   The inner arc collapses toward the center.
'   Shape = outerArc | cap_at_end | line_near_center | cap_at_start
'
' CASE C — Overlapping caps (arc spans nearly 360°):
'   The semicircular end-caps intersect each other. In that case the inner
'   arc is omitted and the two caps are trimmed to their intersection point.
'   Shape = outerArc | trimmedCapEnd | trimmedCapStart
'
' Cap selection is per-end:
'   roundEnd   = True  → semicircular cap at the arc END   point; False → straight radial cap
'   roundStart = True  → semicircular cap at the arc START point; False → straight radial cap
' Case C (caps overlapping, arc near 360°) only applies when BOTH caps are round.
' ---------------------------------------------------------------------------
Private Function BuildArcZone(ByVal oEl As Element, _
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
    ' capEnd   begins facing outward (toward ptOuterEnd) and sweeps a half circle toward the inner edge.
    ' capStart begins facing inward  (toward oCenter)    and sweeps a half circle back to the outer edge.
    If roundEnd Then
        Set capEnd = CreateArcElement2(Nothing, ptArcEnd, Dist, Dist, Matrix3dIdentity, _
                                        Point3dPolarAngle(Point3dSubtract(ptOuterEnd, ptArcEnd)), capSweep)
    End If
    If roundStart Then
        Set capStart = CreateArcElement2(Nothing, ptArcStart, Dist, Dist, Matrix3dIdentity, _
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

            Set trimmedCapEnd   = CreateArcElement2(Nothing, ptArcEnd,   Dist, Dist, Matrix3dIdentity, _
                                                     angCES, Geometry.NormalizeAngle(angCEE - angCES, capSweep))
            Set trimmedCapStart = CreateArcElement2(Nothing, ptArcStart, Dist, Dist, Matrix3dIdentity, _
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

' ============================================================
'  OUTPUT HELPERS
' ============================================================

' FuseRegions
' ---------------------------------------------------------------------------
' Fuses a set of region elements into clean merged outline(s) and returns them in outEls()/
' nOutEls (0-based, nOutEls = count). Shared by every dispatcher that accumulates per-piece
' buffers (lines, stadiums, arc sectors) and needs a single union.
'
' GetRegionUnion is unreliable at large DGN coordinates (a MicroStation precision bug), so every
' buffer is first translated near the origin, unioned, then each result is translated back.
'   - nBuf <= 0 → no output.
'   - nBuf  = 1 → the single buffer is returned as-is (no union needed).
' NOTE: the input buffers are moved in place (near origin) as part of the workaround; callers
' must not reuse bufs() afterwards.
' ---------------------------------------------------------------------------
Private Sub FuseRegions(ByRef bufs() As Element, _
                        ByVal nBuf As Long, _
                        ByRef outEls() As Element, _
                        ByRef nOutEls As Long)
    On Error GoTo ErrorHandler
    nOutEls = 0
    If nBuf <= 0 Then Exit Sub

    If nBuf = 1 Then
        ReDim outEls(0 To 0)
        Set outEls(0) = bufs(0)
        nOutEls = 1
        Exit Sub
    End If

    ' Translate near origin (precision workaround), keeping the inverse offset to restore later.
    Dim toOrigin   As Point3d
    Dim fromOrigin As Point3d
    Dim k          As Long
    toOrigin   = Point3dNegate(bufs(0).Range.High)
    fromOrigin = Point3dNegate(toOrigin)
    For k = 0 To nBuf - 1
        bufs(k).Move toOrigin
    Next k

    ' GetRegionUnion expects region1 = a 1-element array (first shape) and region2 = the rest.
    Dim region1(0 To 0) As Element
    Set region1(0) = bufs(0)
    Dim region2() As Element
    ReDim region2(0 To nBuf - 2)
    For k = 1 To nBuf - 1
        Set region2(k - 1) = bufs(k)
    Next k

    Dim oEnum As ElementEnumerator
    Set oEnum = GetRegionUnion(region1, region2, Nothing, msdFillModeNotFilled)
    If oEnum Is Nothing Then Exit Sub

    Dim resEl As Element
    Do While oEnum.MoveNext
        Set resEl = oEnum.Current
        resEl.Move fromOrigin                 ' restore to the original location
        ReDim Preserve outEls(0 To nOutEls)
        Set outEls(nOutEls) = resEl
        nOutEls = nOutEls + 1
    Loop
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.FuseRegions"
End Sub

' WriteDebugClones
' ---------------------------------------------------------------------------
' Writes a clone of each pre-merge buffer to the model so the individual zones are visible
' alongside the final merged result (DebugMode only).
' ---------------------------------------------------------------------------
Private Sub WriteDebugClones(ByRef bufs() As Element, _
                             ByVal nBuf As Long, _
                             ByVal TargetLevel As Level, _
                             ByVal Color As Long, _
                             ByVal Style As String, _
                             ByVal Weight As Long)
    On Error GoTo ErrorHandler
    Dim k As Long
    For k = 0 To nBuf - 1
        WriteEl bufs(k).Clone, TargetLevel, Color, Style, Weight
    Next k
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.WriteDebugClones"
End Sub

' AddOrWrite
' ---------------------------------------------------------------------------
' Central routing helper called by every dispatcher after building a zone.
'
' The nOut parameter acts as a sentinel to select the write strategy:
'   nOut < 0  → write the element directly to the model right now.
'               Used when MergeZones = False (no merging required).
'   nOut >= 0 → append the element to outBufs() and increment nOut.
'               The caller (Zoning) will later fuse all buffered zones with
'               GetRegionUnion and write the merged result.
' ---------------------------------------------------------------------------
Private Sub AddOrWrite(ByVal oEl As Element, _
                       ByVal TargetLevel As Level, _
                       ByVal Color As Long, _
                       ByVal Style As String, _
                       ByVal Weight As Long, _
                       ByRef outBufs() As Element, _
                       ByRef nOut As Long)
    On Error GoTo ErrorHandler
    If nOut < 0 Then
        WriteEl oEl, TargetLevel, Color, Style, Weight
    Else
        ReDim Preserve outBufs(0 To nOut)
        Set outBufs(nOut) = oEl
        nOut = nOut + 1
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.AddOrWrite"
End Sub

' WriteEl
' Applies symbology and adds the element to the active model.
' This is the only place in this module where elements are written.
Private Sub WriteEl(ByVal oElement As Element, _
                    ByVal TargetLevel As Level, _
                    ByVal Color As Long, _
                    ByVal Style As String, _
                    ByVal Weight As Long)
    On Error GoTo ErrorHandler
    ApplySym oElement, TargetLevel, Color, Style, Weight
    ActiveModelReference.AddElement oElement
    Exit Sub
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.WriteEl"
End Sub

' ApplySym
' Applies level, color, line style, and line weight to an element.
' Parameters equal to -1 or "" are left at the model default.
Private Sub ApplySym(ByVal oEl As Element, _
                     ByVal TargetLevel As Level, _
                     ByVal Color As Long, _
                     ByVal Style As String, _
                     ByVal Weight As Long)
    On Error GoTo ErrorHandler
    oEl.Level = TargetLevel
    If Color  >= 0    Then oEl.Color      = Color
    If Weight >= 0    Then oEl.LineWeight = Weight
    If Len(Style) > 0 Then oEl.LineStyle  = ActiveDesignFile.LineStyles.Find(Style)
    Exit Sub
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.ApplySym"
End Sub