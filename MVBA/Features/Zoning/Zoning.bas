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

' Fusion trace (DebugMode only): the file keeps every line, the Immediate window shows the first ones.
Private Const DBG_FILE As String = "C:\ARES\ARES_zoning_debug.log"
Private Const DBG_ECHO_MAX As Long = 120
Private mnDbgShown As Long

' Round end-caps are built a hair wider than the offset they close, as a RATIO of that offset.
'
' At exactly Dist the cap circle is TANGENT to the two offset lines and to the neighbouring buffer's
' flank: boundaries that coincide exactly, which is the case every boolean engine handles worst.
' GetRegionUnion was seen to drop such a cap outright and with it the end of a zone, to split a
' merged zone in two, and to leave a cap circle visible inside the result. All three went away at
' once when the cap was widened.
'
' 0.0005 is 1 mm at the 2 m zoning distance - the value measured to work - and scales from there, so
' a 0.2 m outline gets 0.1 mm rather than the same millimetre. Worth knowing at small distances: what
' the overlap really has to clear is the file's STORAGE RESOLUTION, since MicroStation keeps
' coordinates as whole UORs and a difference below the resolution rounds back onto the same UOR,
' leaving the boundaries exactly coincident and the bug with them. Check ActiveModelReference's
' resolution before trusting this at a distance far below a metre.
Private Const CAP_OVERLAP_RATIO As Double = 0.0005

' Contour cleanup, one switch per pass - they were tested together and did not behave the same way.
'
' Thinning micro-segments is ON: it only ever removes INTERIOR vertices of a linestring, so the
' endpoints that join it to its neighbours cannot move and the contour cannot open. Nothing it does
' can be undone by the next junction.
'
' Dropping flat arcs is ON again, now that the criterion also bounds the arc's LENGTH. It was turned
' off after behaving on one drawing and misbehaving on the next, and the angle-only test is the likely
' reason: it selected long, gently curved arcs, so closing their gap dragged a junction across a real
' distance. With the length capped at a share of the offset, the endpoint being moved travels at most
' the sliver's own chord. It still touches a junction rather than the inside of a single edge, which
' is the riskier operation of the two - if zones start deforming again, this is the switch to flip.
Private Const ENABLE_VERTEX_THINNING As Boolean = True
Private Const ENABLE_FLAT_ARC_DROP As Boolean = True

' The floor for a straight piece of contour, as a SHARE of the offset distance. It governs both the
' interior vertices of a linestring (CleanTinyVertices) and a straight edge standing on its own in
' the chain (DropSliverEdges) - the same length has to meet the same fate wherever it sits.
'
' 0.01 is 2 cm at the 2 m zoning distance. It was 0.005, and a measured run showed exactly what that
' left behind: vertices at 0.0103 to 0.0194 m and Lines at 0.0116 and 0.0136 m, all sitting just the
' wrong side of a 1 cm floor. 2 cm clears that family and is still a hundredth of the zone's width.
Private Const VERTEX_MERGE_RATIO As Double = 0.01

' An arc is dropped from a merged contour only when it is BOTH nearly flat AND short - see
' DropSliverEdges. Neither test works alone: the angle alone straightens a long, gentle cable curve
' (length is radius x sweep, so a few degrees on a large radius is a real bend); the length alone
' would flatten a genuinely tight little corner.
'
' The length is a SHARE of the offset distance - 0.125 is 0.25 m at the 2 m zoning distance - so it
' follows the offset instead of having to be re-picked.
'
' The angle is what these two constants have to be read TOGETHER for. The slivers come from the cap
' circles, whose radius is the offset distance itself, so a sliver at the length limit sweeps
' FLAT_ARC_LEN_RATIO radians - 7.16 degrees at 0.125. An angle limit below that cuts before the
' length ever binds, and 6 was below it: a measured run kept ten cap slivers of 0.21 to 0.25 m
' purely on 6.1 to 7.1 degrees. At 10 the length is the binding test on anything as round as a cap,
' and the angle is left doing the only job it is good at - refusing to straighten a tight corner,
' which at the length limit means a radius under 1.43 m.
Private Const FLAT_ARC_DEG As Double = 10#
Private Const FLAT_ARC_LEN_RATIO As Double = 0.125

' A part of a cell smaller than this SHARE of the offset distance SQUARED is dropped from the zone -
' 0.25 gives 1 m2 at the 2 m zoning distance. A share again, not a fixed area, so it follows the
' zoning distance instead of having to be re-picked; being an area it goes with the square.
'
' The biggest part is never dropped, whatever its size. Everything else in a cell is a crumb the
' union left behind - a hole barely wider than the overlap, a scrap of outline - but the biggest one
' IS the zone, and a zone smaller than the threshold has to survive it.
Private Const MIN_CELL_PART_AREA_RATIO As Double = 0.25

' How many times the sliver pass may sweep one contour. It judges a sliver by its NEIGHBOURS, and it
' judges them on the chain as it stands - so a sliver between two arcs is refused even when both of
' those arcs are dropped by the very same sweep, which is what a measured run showed happening to
' every straight sliver without exception. The next sweep sees the rebuilt chain, where the arcs are
' gone and the neighbours are straight, and takes it. Sweeping stops as soon as one changes nothing.
Private Const SLIVER_SWEEPS As Long = 4


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

    ' Each run starts its own echo budget and its own block in the trace file.
    mnDbgShown = 0
    If DebugMode Then DbgLine "=== zoning run " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " ==="

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
        FuseRegions allBufs, nAllBufs, mergedAll, nMergedAll, DebugMode, "global"
        For k = 0 To nMergedAll - 1
            WriteEl mergedAll(k), TargetLevel, Color, Style, Weight, Dist
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
    If Not elem Is Nothing Then AddOrWrite elem, TargetLevel, Color, Style, Weight, outBufs, nOut, Dist
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

' DbgLine
' ---------------------------------------------------------------------------
' Debug output for the fusion trace. Goes to a FILE first and to the Immediate window only for the
' first lines of each run: a drawing that folds a thousand buffers floods the Immediate window's
' buffer and pushes the beginning - where the useful part is - out of reach.
' ---------------------------------------------------------------------------
Private Sub DbgLine(ByVal sMsg As String)
    On Error Resume Next

    Dim f As Integer
    f = FreeFile
    Open DBG_FILE For Append As #f
    Print #f, Format(Now, "hh:nn:ss") & "  " & sMsg
    Close #f

    mnDbgShown = mnDbgShown + 1
    If mnDbgShown <= DBG_ECHO_MAX Then
        Debug.Print sMsg
    ElseIf mnDbgShown = DBG_ECHO_MAX + 1 Then
        Debug.Print "... rest of the trace in " & DBG_FILE
    End If
End Sub

' FuseRegions
' ---------------------------------------------------------------------------
' Fuses a set of region elements into clean merged outline(s) and returns them in outEls()/
' nOutEls (0-based, nOutEls = count). Shared by every dispatcher that accumulates per-piece
' buffers (lines, stadiums, arc sectors) and needs a single union.
'
' GetRegionUnion is unreliable at large DGN coordinates (a MicroStation precision bug), so every
' buffer is first translated near the origin, unioned, then each result is translated back.
'   - nBuf <= 0 -> no output.
'   - nBuf  = 1 -> the single buffer is returned as-is (no union needed).
' NOTE: the input buffers are moved in place (near origin) as part of the workaround; callers
' must not reuse bufs() afterwards.
'
' Three guards once lived here - fall back to folding the buffers one at a time when the union
' returned nothing, returned MORE shapes than it was given, or left a buffer inside none of its
' results. All three are gone: they treated symptoms of a union that misbehaves on exactly
' coincident boundaries, fixed nothing, and broke a zone that had been merging correctly. The cause
' is upstream, in the buffer geometry itself. Git history has them if the question reopens.
' ---------------------------------------------------------------------------
Private Sub FuseRegions(ByRef bufs() As Element, _
                        ByVal nBuf As Long, _
                        ByRef outEls() As Element, _
                        ByRef nOutEls As Long, _
                        Optional ByVal DebugMode As Boolean = False, _
                        Optional ByVal sWhere As String = "")
    On Error GoTo ErrorHandler
    nOutEls = 0

    ' Announced on ENTRY as well as on exit: a call that fails never reaches its closing line, so
    ' without this the guilty one is simply absent from the log and cannot be told from a call that
    ' never happened.
    If DebugMode Then DbgLine "FUSE " & sWhere & " : " & nBuf & " in ..."
    If nBuf <= 0 Then Exit Sub

    If nBuf = 1 Then
        If DebugMode Then DbgLine "FUSE " & sWhere & " : 1 in -> 1 out (no union needed)"
        ReDim outEls(0 To 0)
        Set outEls(0) = bufs(0)
        nOutEls = 1
        Exit Sub
    End If

    ' Translate near origin (precision workaround), keeping the inverse offset to restore later.
    Dim toOrigin   As Point3d
    Dim fromOrigin As Point3d
    Dim k          As Long
    toOrigin = Point3dNegate(bufs(0).Range.High)
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
    If oEnum Is Nothing Then
        If DebugMode Then DbgLine "FUSE " & sWhere & " : " & nBuf & " in -> NOTHING (GetRegionUnion returned no enumerator)"
        Exit Sub
    End If

    Dim resEl As Element
    Do While oEnum.MoveNext
        Set resEl = oEnum.Current
        resEl.Move fromOrigin                 ' restore to the original location
        ReDim Preserve outEls(0 To nOutEls)
        Set outEls(nOutEls) = resEl
        nOutEls = nOutEls + 1
    Loop

    If DebugMode Then DbgLine "FUSE " & sWhere & " : " & nBuf & " in -> " & nOutEls & " out"
    Exit Sub

ErrorHandler:
    ' Name the call and how far it got: the caller keeps whatever was collected before the failure,
    ' so a partial result here is a silently incomplete zoning.
    ErrorHandler.HandleError sWhere & " - " & nBuf & " in, " & nOutEls & " collected when it failed - " & _
                             Err.Description, Err.Number, Err.Source, "Zoning.FuseRegions"
    If DebugMode Then DbgLine "FUSE " & sWhere & " : FAILED after " & nOutEls & " of " & nBuf
End Sub

' CleanTinyVertices
' ---------------------------------------------------------------------------
' Rebuilds a merged shape with the micro-segments removed from INSIDE its linestrings, and returns
' it. Hands the shape back untouched when there is nothing to do or anything goes wrong: this is a
' cosmetic pass and must never cost a zone.
'
' Where they come from: the round caps are built CAP_OVERLAP wider than the offset they close (see
' the constant), which is what stopped the union misbehaving on exactly coincident boundaries. The
' overlap survives in the result as 1 mm zigzags along the contour.
'
' Two deliberate limits, both learned the hard way on this module:
'   - ONLY interior vertices, never the first or the last. The endpoints are what join a linestring
'     to its neighbours in the chain, so leaving them alone means the contour cannot open and no
'     junction has to be stitched.
'   - ARCS ARE NOT TOUCHED, at all. StartPoint/EndPoint are writable on an ArcElement but writing
'     them does not move the arc: it re-solves it through the new point, and radius and sweep go
'     wild. An earlier attempt did exactly that and produced arcs looping over themselves.
'
' RemoveVertex is ZERO-based, unlike most indexes in this object library - the docs say so
' explicitly. Vertices are dropped highest-index-first so the lower ones stay valid.
' ---------------------------------------------------------------------------
Private Function CleanTinyVertices(ByVal oShape As Element, ByVal dTol As Double) As Element
    On Error GoTo ErrorHandler

    Set CleanTinyVertices = oShape
    If oShape Is Nothing Then Exit Function
    If oShape.Type <> msdElementTypeComplexShape Then Exit Function

    Dim subs()  As Element
    Dim nSub    As Long
    Dim oEE     As ElementEnumerator
    Set oEE = oShape.AsComplexShapeElement.GetSubElements
    nSub = 0
    Do While oEE.MoveNext
        ReDim Preserve subs(0 To nSub)
        Set subs(nSub) = oEE.Current
        nSub = nSub + 1
    Loop
    If nSub < 2 Then Exit Function

    Dim i        As Long
    Dim nDropped As Long
    nDropped = 0
    For i = 0 To nSub - 1
        If subs(i).Type = msdElementTypeLineString Then
            nDropped = nDropped + ThinLineString(subs(i), dTol)
        End If
    Next i
    If nDropped = 0 Then Exit Function       ' nothing to gain: hand back the original

    Dim chain() As ChainableElement
    ReDim chain(0 To nSub - 1)
    For i = 0 To nSub - 1
        Set chain(i) = subs(i)
    Next i

    Dim oNew As ComplexShapeElement
    Set oNew = CreateComplexShapeElement1(chain, msdFillModeNotFilled)
    If Not oNew Is Nothing Then Set CleanTinyVertices = oNew
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.CleanTinyVertices"
    Set CleanTinyVertices = oShape
End Function

' ThinLineString
' ---------------------------------------------------------------------------
' Drops the interior vertices that sit closer than dTol to the last one kept, and returns how many
' went. The first and last vertices are never candidates. Modifies oEl in place; the caller rebuilds
' the parent shape around it.
' ---------------------------------------------------------------------------
Private Function ThinLineString(ByVal oEl As Element, ByVal dTol As Double) As Long
    On Error GoTo ErrorHandler

    ThinLineString = 0
    If Not oEl.IsVertexList Then Exit Function

    Dim oVL    As VertexList
    Set oVL = oEl
    Dim verts() As Point3d
    verts = oVL.GetVertices
    Dim nV As Long
    nV = UBound(verts) - LBound(verts) + 1

    ' Two vertices is a plain segment with no interior to thin. It is not a dead end though: a
    ' two-vertex linestring is a straight edge, and DropSliverEdges takes it from there.
    If nV < 3 Then Exit Function

    ' Forward pass against the last KEPT vertex, so a run of micro-segments collapses to one point
    ' instead of surviving as a chain of pairs each just under the tolerance.
    Dim drop()  As Long
    Dim nDrop   As Long
    Dim iLast   As Long
    Dim i       As Long
    ReDim drop(0 To nV - 1)
    nDrop = 0
    iLast = LBound(verts)

    For i = LBound(verts) + 1 To UBound(verts) - 1
        If Point3dDistance(verts(i), verts(iLast)) < dTol Then
            drop(nDrop) = i - LBound(verts)   ' RemoveVertex is zero-based
            nDrop = nDrop + 1
        Else
            iLast = i
        End If
    Next i

    ' The LAST segment needs its own test, and it is the one that was being missed. Walking forward
    ' only ever measures a vertex against what precedes it, so a short final segment - the last kept
    ' vertex sitting a millimetre from the endpoint - is never seen: the endpoint is excluded as a
    ' junction, and its predecessor looked fine when measured from the other side.
    ' The fix keeps the endpoint exactly where it is and drops its PREDECESSOR instead, so the
    ' contour runs straight into the junction and the short segment is not rebuilt.
    If iLast > LBound(verts) Then
        If Point3dDistance(verts(UBound(verts)), verts(iLast)) < dTol Then
            drop(nDrop) = iLast - LBound(verts)
            nDrop = nDrop + 1
        End If
    End If

    If nDrop = 0 Then Exit Function
    ' Every vertex a candidate means a linestring that is nothing but micro-segments. Thinning it
    ' would leave fewer than the two endpoints, so it survives whole and DropSliverEdges gets it.
    If nV - nDrop < 2 Then Exit Function

    ' Highest index first: removing a low one would shift every index after it. The closing test
    ' appends an index that is lower than the ones before it, so sort before removing.
    Dim j As Long
    Dim t As Long
    For i = 0 To nDrop - 2
        For j = 0 To nDrop - 2 - i
            If drop(j) < drop(j + 1) Then
                t = drop(j) : drop(j) = drop(j + 1) : drop(j + 1) = t
            End If
        Next j
    Next i

    For i = 0 To nDrop - 1
        oVL.RemoveVertex drop(i)
    Next i

    ThinLineString = nDrop
    Exit Function

ErrorHandler:
    ThinLineString = 0
End Function

' DropSliverEdges
' ---------------------------------------------------------------------------
' Rebuilds a merged shape without its sliver edges and returns it. Hands the shape back untouched
' when there is nothing to do or anything goes wrong: cosmetic pass, it must never cost a zone.
'
' Two kinds of sliver, one stitching:
'   - an ARC that is BOTH nearly flat AND short. The sweep says it carries no curvature worth
'     keeping; the length says it is a sliver and not a gentle bend spread over a large radius. The
'     angle alone would damage real geometry - length is radius x sweep, so a few degrees on a big
'     radius is a long, deliberate cable curve, and dropping it would straighten a bend the drawing
'     meant to have.
'   - a STRAIGHT edge shorter than the thinning tolerance. The same length inside a linestring is a
'     vertex CleanTinyVertices drops without hesitation; standing on its own in the chain there was
'     nothing at all to remove it, and a measured run found 1 cm Lines outliving every pass. One
'     tolerance, one outcome, wherever the segment happens to sit.
'
' Two guards make the stitching safe:
'   - ONE straight neighbour (Line or LineString) is enough, and it is that side which moves, onto
'     the vanishing arc's own far end. Between two ARCS the sliver is left alone: closing that gap
'     would mean moving an arc's endpoint, and writing StartPoint/EndPoint on an ArcElement
'     re-solves it through the new point - radius and sweep go wild, which is how an earlier attempt
'     produced arcs looping over themselves.
'   - exactly one side moves, so no junction is ever pulled from both ends. The previous neighbour
'     is preferred; the next one is used when the previous is an arc. A Line is recreated (its
'     endpoints are read-only), a LineString has the relevant end vertex modified, ModifyVertex
'     being zero-based like RemoveVertex.
'
' The gap being closed is the vanishing edge's own chord, which is why both criteria are size
' criteria: nothing is ever dropped that would move a junction further than the sliver was long.
' ---------------------------------------------------------------------------
Private Function DropSliverEdges(ByVal oShape As Element, _
                              ByVal dMaxDeg As Double, _
                              ByVal dMaxLen As Double, _
                              ByVal dMinEdge As Double) As Element
    On Error GoTo ErrorHandler

    Set DropSliverEdges = oShape
    If oShape Is Nothing Then Exit Function
    If oShape.Type <> msdElementTypeComplexShape Then Exit Function

    Dim subs()  As Element
    Dim nSub    As Long
    Dim oEE     As ElementEnumerator
    Set oEE = oShape.AsComplexShapeElement.GetSubElements
    nSub = 0
    Do While oEE.MoveNext
        ReDim Preserve subs(0 To nSub)
        Set subs(nSub) = oEE.Current
        nSub = nSub + 1
    Loop
    If nSub < 4 Then Exit Function

    Dim bDrop() As Boolean
    ReDim bDrop(0 To nSub - 1)

    Dim i      As Long
    Dim iPrev  As Long
    Dim iNext  As Long
    Dim nGone  As Long
    Dim dDeg   As Double
    Dim dLen   As Double
    Dim bGo    As Boolean
    nGone = 0

    For i = 0 To nSub - 1
        bGo = False

        Select Case subs(i).Type

            Case msdElementTypeArc
                dDeg = Abs(subs(i).AsArcElement.SweepAngle) * 180# / Application.PI
                dLen = subs(i).AsArcElement.Length

                ' Nearly flat AND short. The sweep says the arc carries no curvature worth keeping;
                ' the length says it is a sliver, not a gentle bend spread over a large radius.
                bGo = (dDeg <= dMaxDeg And dLen <= dMaxLen)

            Case msdElementTypeLine, msdElementTypeLineString
                ' A straight edge shorter than the thinning tolerance. The SAME length inside a
                ' linestring is a vertex CleanTinyVertices removes without a second thought; standing
                ' on its own in the chain there was nothing at all to remove it, which is why 1 cm
                ' segments outlived every pass. One tolerance, one outcome, wherever the segment sits.
                dLen = StraightSliver(subs(i))
                bGo = (dLen >= 0# And dLen < dMinEdge)

        End Select

        If bGo Then
            iPrev = (i + nSub - 1) Mod nSub
            iNext = (i + 1) Mod nSub

            ' ONE straight neighbour is enough, and that is the side that moves. What must never
            ' happen is moving an ARC's endpoint: writing StartPoint/EndPoint on an ArcElement
            ' re-solves it through the new point and its radius and sweep go wild. So the gap is
            ' always closed from the straight side, onto the vanishing edge's own far end.
            If IsStraight(subs(iPrev)) And Not bDrop(iPrev) Then
                Set subs(iPrev) = ExtendTo(subs(iPrev), subs(i).AsChainableElement.EndPoint)
                bDrop(i) = True
                nGone = nGone + 1
            ElseIf IsStraight(subs(iNext)) And Not bDrop(iNext) Then
                Set subs(iNext) = StartFrom(subs(iNext), subs(i).AsChainableElement.StartPoint)
                bDrop(i) = True
                nGone = nGone + 1
            End If
        End If
    Next i

    If nGone = 0 Then Exit Function
    If nSub - nGone < 3 Then Exit Function

    Dim chain() As ChainableElement
    Dim nKeep   As Long
    ReDim chain(0 To nSub - nGone - 1)
    nKeep = 0
    For i = 0 To nSub - 1
        If Not bDrop(i) Then
            Set chain(nKeep) = subs(i)
            nKeep = nKeep + 1
        End If
    Next i

    Dim oNew As ComplexShapeElement
    Set oNew = CreateComplexShapeElement1(chain, msdFillModeNotFilled)
    If Not oNew Is Nothing Then Set DropSliverEdges = oNew
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.DropSliverEdges"
    Set DropSliverEdges = oShape
End Function

' StraightSliver
' ---------------------------------------------------------------------------
' The length of a contour edge that is straight from end to end, or -1 for anything else.
'
' A Line qualifies outright. A LineString qualifies only with exactly TWO vertices, which makes it
' the same thing geometrically - a longer one has interior vertices, and those belong to
' CleanTinyVertices, which thins them without touching either junction.
' ---------------------------------------------------------------------------
Private Function StraightSliver(ByVal oEl As Element) As Double
    On Error GoTo ErrorHandler

    StraightSliver = -1#

    If oEl.Type = msdElementTypeLine Then
        StraightSliver = Point3dDistance(oEl.AsLineElement.StartPoint, oEl.AsLineElement.EndPoint)
    ElseIf oEl.Type = msdElementTypeLineString Then
        Dim oVL As VertexList
        Set oVL = oEl
        If oVL.VerticesCount = 2 Then _
            StraightSliver = Point3dDistance(oEl.AsChainableElement.StartPoint, oEl.AsChainableElement.EndPoint)
    End If
    Exit Function

ErrorHandler:
    StraightSliver = -1#        ' unmeasurable: never a candidate
End Function

' True for the contour edges whose end can be moved safely - the straight ones.
Private Function IsStraight(ByVal oEl As Element) As Boolean
    IsStraight = (oEl.Type = msdElementTypeLine Or oEl.Type = msdElementTypeLineString)
End Function

' ExtendTo
' ---------------------------------------------------------------------------
' Returns oEl ending at ptEnd. A Line is recreated, its endpoints being read-only; a LineString has
' its LAST vertex modified in place. Only ever called on a straight edge - see IsStraight.
' ---------------------------------------------------------------------------
Private Function ExtendTo(ByVal oEl As Element, ByRef ptEnd As Point3d) As Element
    On Error GoTo ErrorHandler

    Set ExtendTo = oEl

    Select Case oEl.Type
        Case msdElementTypeLine
            Set ExtendTo = CreateLineElement2(Nothing, oEl.AsLineElement.StartPoint, ptEnd)
        Case msdElementTypeLineString
            Dim oVL As VertexList
            Set oVL = oEl
            oVL.ModifyVertex oVL.VerticesCount - 1, ptEnd     ' zero-based, like RemoveVertex
            Set ExtendTo = oEl
    End Select
    Exit Function

ErrorHandler:
    Set ExtendTo = oEl
End Function

' StartFrom
' ---------------------------------------------------------------------------
' Returns oEl starting at ptStart - the mirror of ExtendTo, for when the straight neighbour is the
' one AFTER the arc being removed. A Line is recreated, its endpoints being read-only; a LineString
' has its FIRST vertex modified. Only ever called on a straight edge - see IsStraight.
' ---------------------------------------------------------------------------
Private Function StartFrom(ByVal oEl As Element, ByRef ptStart As Point3d) As Element
    On Error GoTo ErrorHandler

    Set StartFrom = oEl

    Select Case oEl.Type
        Case msdElementTypeLine
            Set StartFrom = CreateLineElement2(Nothing, ptStart, oEl.AsLineElement.EndPoint)
        Case msdElementTypeLineString
            Dim oVL As VertexList
            Set oVL = oEl
            oVL.ModifyVertex 0, ptStart                       ' zero-based, like RemoveVertex
            Set StartFrom = oEl
    End Select
    Exit Function

ErrorHandler:
    Set StartFrom = oEl
End Function

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
                       ByRef nOut As Long, _
                       Optional ByVal Dist As Double = 0)
    On Error GoTo ErrorHandler
    If nOut < 0 Then
        WriteEl oEl, TargetLevel, Color, Style, Weight, Dist
    Else
        ReDim Preserve outBufs(0 To nOut)
        Set outBufs(nOut) = oEl
        nOut = nOut + 1
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.AddOrWrite"
End Sub

' UnwrapLoneCell
' ---------------------------------------------------------------------------
' Returns the single part of a cell that has only one left, and the cell untouched otherwise.
'
' A zone with a hole arrives as a cell grouping the outline and its island(s). Once the size floor
' has taken the crumbs out there is often nothing left but the outline, and a cell wrapped around one
' shape groups nothing: it just makes the zone awkward to select and to measure.
'
' Deleting the cell costs nothing here because it was never written - WriteEl adds the element it is
' handed, so returning the part instead of the cell IS the deletion. Do not move this anywhere the
' cell already lives in the file.
' ---------------------------------------------------------------------------
Private Function UnwrapLoneCell(ByVal oCell As Element) As Element
    On Error GoTo ErrorHandler

    Set UnwrapLoneCell = oCell

    Dim oEE    As ElementEnumerator
    Dim oOnly  As Element
    Dim nParts As Long

    Set oEE = oCell.AsCellElement.GetSubElements
    Do While oEE.MoveNext
        nParts = nParts + 1
        If nParts > 1 Then Exit Function          ' still grouping something: the cell stays
        Set oOnly = oEE.Current
    Loop
    If nParts <> 1 Then Exit Function
    If oOnly Is Nothing Then Exit Function

    Set UnwrapLoneCell = oOnly
    Exit Function

ErrorHandler:
    Set UnwrapLoneCell = oCell
End Function

' CleanContour
' ---------------------------------------------------------------------------
' Runs whichever cleanup passes are enabled over ONE closed contour and returns the result. Both
' passes hand the element back untouched when they cannot help, so this is safe on anything.
' ---------------------------------------------------------------------------
Private Function CleanContour(ByVal oEl As Element, ByVal Dist As Double) As Element
    Dim oOut As Element
    Dim oWas As Element
    Dim k    As Long

    Set oOut = oEl
    If ENABLE_VERTEX_THINNING Then Set oOut = CleanTinyVertices(oOut, Dist * VERTEX_MERGE_RATIO)

    ' Swept until it settles - see SLIVER_SWEEPS. Dropping a sliver frees its neighbours to become
    ' droppable in turn, and one sweep can only ever see the chain it started with.
    If ENABLE_FLAT_ARC_DROP Then
        For k = 1 To SLIVER_SWEEPS
            Set oWas = oOut
            Set oOut = DropSliverEdges(oOut, FLAT_ARC_DEG, Dist * FLAT_ARC_LEN_RATIO, _
                                       Dist * VERTEX_MERGE_RATIO)
            ' The pass hands back the SAME object when it dropped nothing: that is the fixed point.
            If oOut Is oWas Then Exit For
        Next k
    End If

    Set CleanContour = oOut
End Function

' CleanCellChildren
' ---------------------------------------------------------------------------
' Cleans the contours held INSIDE a cell, in place. A zone with a hole is returned by the union as a
' cell grouping two or more complex shapes - the outline and its island(s) - so it used to reach
' WriteEl as element type 2 and both passes declined it on the spot, leaving those zones with every
' sliver they were born with.
'
' The cell is walked with its OWN cursor - ResetElementEnumeration, MoveToNextElement,
' CopyCurrentElement, ReplaceCurrentElement - and never rebuilt from its children: there is no API
' here that recreates a grouped hole, and CreateCellElement would give back a plain cell whose
' island renders as a second solid instead of a hole.
'
' Two things happen to each part. A part under MIN_CELL_PART_AREA_RATIO is DELETED outright - the
' crumbs the union leaves inside a cell, holes barely wider than the cap overlap - except the
' biggest part, which is the zone itself and is never dropped whatever its size. Everything that
' survives is cleaned like any other contour. DeleteCurrentElement leaves the marker on the element
' BEFORE the one it removed, so deleting inside the walk is safe and skips nothing.
'
' StepThroughNestingChanges is False. A nested cell is not something this module produces, and
' stepping into one would hand back children this code has no business rewriting.
'
' The hole flag cannot be written onto a finished element - IsHole is READ-ONLY on a closed element.
' A shape rebuilt by CreateComplexShapeElement1 takes the ACTIVE area mode instead, so the active
' mode is set to match each child BEFORE cleaning it and restored on the way out. A first run without
' this replaced the outline and refused all three holes of a zone, which is what the check below was
' there to catch.
'
' The check stays regardless: the replacement is compared with the original and abandoned when the
' flag differs. A tidier contour is never worth turning a hole into a solid.
' ---------------------------------------------------------------------------
Private Sub CleanCellChildren(ByVal oCell As CellElement, ByVal Dist As Double)
    On Error GoTo ErrorHandler

    Dim oChild   As Element
    Dim oNew     As Element
    Dim nSeen    As Long
    Dim nDone    As Long
    Dim nGone    As Long
    Dim bHole    As Boolean
    Dim bRestore As Boolean
    Dim dMinArea As Double
    Dim dBiggest As Double
    Dim dArea    As Double

    dMinArea = Dist * Dist * MIN_CELL_PART_AREA_RATIO
    dBiggest = BiggestPart(oCell)
    bRestore = ActiveAreaHole

    oCell.ResetElementEnumeration
    Do While oCell.MoveToNextElement(False)
        Set oChild = oCell.CopyCurrentElement
        If IsAreaPart(oChild) Then
            nSeen = nSeen + 1
            dArea = AreaOf(oChild)

            If dArea < dMinArea And dArea < dBiggest Then
                ' A crumb: too small to mean anything, and not the part that carries the zone.
                oCell.DeleteCurrentElement
                nGone = nGone + 1

            ElseIf oChild.Type = msdElementTypeComplexShape Then
                ' The rebuilt shape inherits the ACTIVE area mode, so hand it the child's own.
                bHole = HoleOf(oChild)
                SetActiveAreaHole bHole

                Set oNew = CleanContour(oChild, Dist)
                If Not oNew Is oChild Then
                    If SameHoleFlag(oChild, oNew) Then
                        oCell.ReplaceCurrentElement oNew
                        nDone = nDone + 1
                    End If
                End If
            End If
        End If
    Loop

    SetActiveAreaHole bRestore
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.CleanCellChildren"
End Sub

' True when two closed elements agree on being a hole or a solid. False when the flag cannot be
' read, which keeps the original - see CleanCellChildren.
Private Function SameHoleFlag(ByVal oA As Element, ByVal oB As Element) As Boolean
    On Error GoTo ErrorHandler
    SameHoleFlag = (oA.AsClosedElement.IsHole = oB.AsClosedElement.IsHole)
    Exit Function
ErrorHandler:
    SameHoleFlag = False
End Function

' True for the cell parts that enclose an area, the ones the size floor applies to.
Private Function IsAreaPart(ByVal oEl As Element) As Boolean
    Select Case oEl.Type
        Case msdElementTypeComplexShape, msdElementTypeShape, msdElementTypeEllipse
            IsAreaPart = True
    End Select
End Function

' The area of one closed part. An unreadable area comes back huge, never small: a part is only ever
' deleted on a measurement that succeeded.
Private Function AreaOf(ByVal oEl As Element) As Double
    On Error GoTo ErrorHandler
    AreaOf = oEl.AsClosedElement.Area
    Exit Function
ErrorHandler:
    AreaOf = 1E+30
End Function

' The area of the largest part in a cell - the one the size floor must never remove. Read through
' GetSubElements rather than the cell cursor, which the caller is about to walk itself.
Private Function BiggestPart(ByVal oCell As CellElement) As Double
    On Error GoTo ErrorHandler

    Dim oEE   As ElementEnumerator
    Dim dMax  As Double
    Dim dArea As Double

    Set oEE = oCell.GetSubElements
    Do While oEE.MoveNext
        If IsAreaPart(oEE.Current) Then
            dArea = AreaOf(oEE.Current)
            If dArea > dMax Then dMax = dArea
        End If
    Loop

    BiggestPart = dMax
    Exit Function

ErrorHandler:
    BiggestPart = 0
End Function

' The hole flag of one closed element, False when it cannot be read. Only ever used to choose the
' active area mode; SameHoleFlag is what actually authorises a replacement.
Private Function HoleOf(ByVal oEl As Element) As Boolean
    On Error Resume Next
    HoleOf = oEl.AsClosedElement.IsHole
End Function

' ActiveAreaHole / SetActiveAreaHole
' ---------------------------------------------------------------------------
' The ACTIVE area mode - hole or solid - which is what a newly created shape inherits, there being
' no way to set the flag on a finished element.
'
' Reached LATE, through an Object variable, on purpose: the module must still compile against a
' MicroStation whose Settings object does not expose AreaModeHole. Where it is missing, both calls
' do nothing, the rebuilt holes come back solid, and SameHoleFlag refuses them - the zone keeps the
' contour it had instead of losing its hole.
' ---------------------------------------------------------------------------
Private Function ActiveAreaHole() As Boolean
    On Error Resume Next
    Dim oSettings As Object
    Set oSettings = ActiveSettings
    ActiveAreaHole = oSettings.AreaModeHole
End Function

Private Sub SetActiveAreaHole(ByVal bHole As Boolean)
    On Error Resume Next
    Dim oSettings As Object
    Set oSettings = ActiveSettings
    oSettings.AreaModeHole = bHole
End Sub

' WriteEl
' Applies symbology and adds the element to the active model.
' This is the only place in this module where elements are written.
Private Sub WriteEl(ByVal oElement As Element, _
                    ByVal TargetLevel As Level, _
                    ByVal Color As Long, _
                    ByVal Style As String, _
                    ByVal Weight As Long, _
                    Optional ByVal Dist As Double = 0)
    On Error GoTo ErrorHandler

    ' Contour cleanup happens HERE, on what is actually drawn, and nowhere else. It used to run
    ' inside FuseRegions, which meant every intermediate result was cleaned too: a per-element zone
    ' gets cleaned, then feeds the global fusion, which rebuilds the contour from scratch and throws
    ' that work away. One pass, on the final shape, whatever the depth of merging.
    ' Dist = 0 means "do not touch": that is how the debug clones of pre-merge buffers come through.
    If Dist > 0 Then
        If oElement.Type = msdElementTypeCellHeader Then
            ' A zone that has a hole comes back from the union as a CELL holding the outline and its
            ' island(s), not as a complex shape. Its contours are cleaned inside the cell, and the
            ' cell is dropped when the size floor leaves nothing for it to group.
            CleanCellChildren oElement.AsCellElement, Dist
            Set oElement = UnwrapLoneCell(oElement)
        Else
            Set oElement = CleanContour(oElement, Dist)
        End If
    End If

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