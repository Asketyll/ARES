' Module: RegionSplit
' Description: Geometry engine for the SplitRegion command. Cuts one closed region
'              (Shape / ComplexShape) into two regions with a single datapoint on its
'              boundary. The cut runs perpendicular to the local boundary segment at the
'              clicked point, across the interior to the opposite boundary. Both halves
'              inherit the original's level + symbology; the original is deleted (default)
'              or kept (ARES_RegionSplit_Keep_Original).
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, ARESConstants, ErrorHandlerClass, Geometry, LangManager
'
' Mechanic shipped: a thin "knife" rectangle + GetRegionDifference, evaluated near the
' origin (Zoning precision workaround). The knife half-width derives from
' ARES_RegionSplit_Collinear_Tol (a config var, not a literal).
Option Explicit

' SplitElementAt
' ---------------------------------------------------------------------------
' Sole public engine entry. Validates inputs, computes the edge-perpendicular cut,
' builds and validates the two halves, writes both, then disposes of the original
' per ARES_RegionSplit_Keep_Original.
'
' Ordering guarantee: build + validate both halves FIRST, add both, THEN delete
' the original. Any error before completion leaves the original intact (no destructive
' partial edit). Every degenerate input aborts cleanly (ShowStatus + return, no model
' change) via the input guards.
'
' Parameters:
'   oRegion - the located closed region (Shape / ComplexShape)
'   ClickPt - the clicked datapoint, on / near the region boundary
' Note: ClickPt is a Point3d (a VBA user-defined type) and therefore MUST be passed
' ByRef -- UDTs cannot be passed ByVal in VBA. It is treated as read-only input here.
Public Sub SplitElementAt(ByVal oRegion As Element, ByRef ClickPt As Point3d)
    On Error GoTo ErrorHandler

    Dim dCollinearTol As Double
    Dim dStrokeTol    As Double   ' max chordal deviation when densifying an arc boundary side
    Dim bKeepOriginal As Boolean

    Dim verts()    As Point3d   ' boundary polygon vertices (zero-based, GetVertices)
    Dim nSeg       As Long      ' index of the boundary segment closest to the click
    Dim entryPt    As Point3d   ' perpendicular foot of ClickPt on the closest segment
    Dim dirIn      As Point3d   ' unit-ish interior cut direction (perpendicular to segment)
    Dim exitPt     As Point3d   ' opposite-boundary crossing
    Dim knifeEl    As Element   ' thin rectangle straddling entryPt -> exitPt
    Dim halves()   As Element   ' the two resulting region halves
    Dim nHalves    As Long

    ' --- Read fate / tolerances from config (config vars, not literals) ---
    If Not ReadConfig(dCollinearTol, dStrokeTol, bKeepOriginal) Then
        ShowSplitStatus "RegionSplitCannotSplit", "SplitRegion: configuration unavailable"
        Exit Sub
    End If

    ' --- Validate region + active model ---
    If Not IsSplittableRegion(oRegion) Then
        ShowSplitStatus "RegionSplitNoRegion", "SplitRegion: not a supported closed region"
        ErrorHandler.HandleError "oRegion is Nothing / not a supported closed region", 0, "RegionSplit.SplitElementAt", "WARNING"
        Exit Sub
    End If
    If Not Application.HasActiveModelReference Then
        ShowSplitStatus "RegionSplitCannotSplit", "SplitRegion: no active model"
        ErrorHandler.HandleError "No active model reference", 0, "RegionSplit.SplitElementAt", "WARNING"
        Exit Sub
    End If

    ' --- Resolve boundary vertices and the closest segment ---
    ' Arc-aware: a plain Shape keeps the straight-segment GetVertices path; a ComplexShape
    ' is walked sub-element by sub-element so arc sides are densified into chords (dStrokeTol)
    ' instead of collapsing the vertex list.
    verts = GetBoundaryVertices(oRegion, dStrokeTol, dCollinearTol)
    If Not HasAtLeast(verts, 2) Then
        ShowSplitStatus "RegionSplitCannotSplit", "SplitRegion: cannot read boundary"
        ErrorHandler.HandleError "Boundary vertex list empty / too small", 0, "RegionSplit.SplitElementAt", "WARNING"
        Exit Sub
    End If

    ' Closest boundary segment, computed locally from the resolved vertices. The native
    ' VertexList.GetClosestSegment raised a COM "index not in collection" error on some
    ' region elements (re-casting AsVertexList for the call), so we measure ClickPt against
    ' verts() directly -- deterministic and immune to that COM quirk.
    nSeg = GetClosestSegmentIndex(verts, ClickPt)
    If nSeg < LBound(verts) Or nSeg > UBound(verts) - 1 Then
        ShowSplitStatus "RegionSplitClickNotOnEdge", "SplitRegion: no boundary segment near the click"
        ErrorHandler.HandleError "No closest boundary segment resolved (index " & nSeg & ")", 0, "RegionSplit.SplitElementAt", "WARNING"
        Exit Sub
    End If

    ' Degenerate (zero-length / duplicate vertex) segment guard.
    If Point3dDistanceXY(verts(nSeg), verts(nSeg + 1)) <= dCollinearTol Then
        ShowSplitStatus "RegionSplitClickNotOnEdge", "SplitRegion: clicked segment is degenerate"
        ErrorHandler.HandleError "Closest boundary segment is degenerate (<= collinear tol)", 0, "RegionSplit.SplitElementAt", "WARNING"
        Exit Sub
    End If

    ' --- Entry point: perpendicular foot of the click on the closest segment ---
    If Not GetEntryPoint(verts(nSeg), verts(nSeg + 1), ClickPt, entryPt) Then
        ShowSplitStatus "RegionSplitClickNotOnEdge", "SplitRegion: cannot resolve the entry point"
        ErrorHandler.HandleError "GetEntryPoint failed (degenerate closest segment)", 0, "RegionSplit.SplitElementAt", "WARNING"
        Exit Sub
    End If

    ' --- Interior cut direction: perpendicular to the segment, oriented inward ---
    If Not GetInteriorDirection(verts(nSeg), verts(nSeg + 1), verts, dirIn) Then
        ShowSplitStatus "RegionSplitCannotSplit", "SplitRegion: cannot orient the cut into the interior"
        ErrorHandler.HandleError "Failed to orient the perpendicular into the region interior", 0, "RegionSplit.SplitElementAt", "WARNING"
        Exit Sub
    End If

    ' --- Arc sides: follow the arc's RADIUS (entry -> arc centre) instead of the chord
    '     perpendicular, so the cut runs along the arc's angle toward its origin (user
    '     preference). No-op on straight sides and plain Shapes; non-fatal if it cannot resolve. ---
    RefineDirectionRadialIfArc oRegion, entryPt, dStrokeTol, dCollinearTol, dirIn

    ' --- Exit point: first opposite-boundary crossing beyond the entry ---
    If Not GetExitPoint(oRegion, entryPt, dirIn, dCollinearTol, exitPt) Then
        ShowSplitStatus "RegionSplitCannotSplit", "SplitRegion: cut does not reach the opposite boundary"
        ErrorHandler.HandleError "No valid opposite-boundary crossing for the perpendicular cut", 0, "RegionSplit.SplitElementAt", "WARNING"
        Exit Sub
    End If

    ' --- Build the thin knife straddling entryPt -> exitPt ---
    ' Pass the region's bbox diagonal so the knife width scales with extent (else a large region
    ' collapses the slot below GetRegionDifference's tolerance -> single region).
    Dim oRng As Range3d
    oRng = oRegion.Range
    Set knifeEl = BuildKnife(entryPt, exitPt, dCollinearTol, dStrokeTol, _
                             Point3dDistanceXY(oRng.Low, oRng.High))
    If knifeEl Is Nothing Then
        ShowSplitStatus "RegionSplitCannotSplit", "SplitRegion: failed to build the cut knife"
        ErrorHandler.HandleError "BuildKnife returned Nothing", 0, "RegionSplit.SplitElementAt", "WARNING"
        Exit Sub
    End If

    ' --- Boolean difference near the origin → the two halves ---
    nHalves = 0
    If Not SplitByKnife(oRegion, knifeEl, halves, nHalves) Then
        ShowSplitStatus "RegionSplitCannotSplit", "SplitRegion: boolean split failed"
        ErrorHandler.HandleError "GetRegionDifference failed during split", 0, "RegionSplit.SplitElementAt", "WARNING"
        Exit Sub
    End If

    ' Must yield >= 2 non-empty regions, else abort with no model change.
    If nHalves < 2 Then
        ShowSplitStatus "RegionSplitCannotSplit", "SplitRegion: split did not produce two regions"
        ErrorHandler.HandleError "Boolean split produced fewer than two regions (" & nHalves & ")", 0, "RegionSplit.SplitElementAt", "WARNING"
        Exit Sub
    End If

    ' --- Write both halves with the original's symbology, THEN delete the original ---
    ' RemoveElement is the documented deletion API (Element.DeleteElement is undocumented in
    ' the MVBA reference). It also erases the element from the view.
    ' WriteHalves is a Function returning True ONLY if BOTH halves were added and
    ' styled. The original is deleted ONLY on real success AND Not bKeepOriginal, so a partial
    ' write failure leaves the original intact (anti-destructive ordering holds on the error
    ' path too). On failure: abort cleanly, original untouched, log a WARNING.
    If Not WriteHalves(oRegion, halves, nHalves) Then
        ShowSplitStatus "RegionSplitCannotSplit", "SplitRegion: failed to write both halves"
        ErrorHandler.HandleError "WriteHalves failed; original left intact", 0, "RegionSplit.SplitElementAt", "WARNING"
        Exit Sub
    End If
    If Not bKeepOriginal Then ActiveModelReference.RemoveElement oRegion

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.SplitElementAt"
    ShowSplitStatus "RegionSplitCannotSplit", "SplitRegion failed: " & Err.Description
End Sub

' ============================================================
'  CONFIG / VALIDATION
' ============================================================

' ReadConfig
' Reads the four RegionSplit config vars. Doubles parsed with Val(), the
' Boolean with the UCase(Trim(...)) = "TRUE" idiom (as Command.bas does). Returns
' False on a non-positive / unusable tolerance so the caller aborts cleanly.
' dStrokeTol (ARES_RegionSplit_Stroke_Tol) is the max chordal deviation used to densify
' an arc boundary side into a polyline; it must be > 0 like the others.
Private Function ReadConfig(ByRef dCollinearTol As Double, _
                            ByRef dStrokeTol As Double, _
                            ByRef bKeepOriginal As Boolean) As Boolean
    On Error GoTo ErrorHandler

    ReadConfig = False
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then Exit Function

    dCollinearTol = Val(ARESConfig.ARES_REGIONSPLIT_COLLINEAR_TOL.Value)
    dStrokeTol = Val(ARESConfig.ARES_REGIONSPLIT_STROKE_TOL.Value)
    bKeepOriginal = (UCase(Trim(ARESConfig.ARES_REGIONSPLIT_KEEP_ORIGINAL.Value)) = "TRUE")

    ' Defensive: both tolerances must be strictly positive to be usable.
    If dCollinearTol <= 0 Or dStrokeTol <= 0 Then Exit Function

    ReadConfig = True
    Exit Function

ErrorHandler:
    ReadConfig = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.ReadConfig"
End Function

' IsSplittableRegion
' True only for a supported closed region: Shape or ComplexShape. Ellipse is out of
' scope for this delivery (see implementation note); _LocateFilter also rejects it.
' Defence-in-depth even though _LocateFilter already filtered.
Private Function IsSplittableRegion(ByVal oRegion As Element) As Boolean
    On Error GoTo ErrorHandler
    IsSplittableRegion = False
    If oRegion Is Nothing Then Exit Function
    If Not oRegion.IsGraphical Then Exit Function
    IsSplittableRegion = (oRegion.IsShapeElement Or oRegion.IsComplexShapeElement)
    Exit Function

ErrorHandler:
    IsSplittableRegion = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.IsSplittableRegion"
End Function

' GetBoundaryVertices
' Resolves the closed boundary as a polyline the cut pipeline can consume.
' Arc-aware: a ComplexShape is walked sub-element by sub-element so arc
' sides are densified into chords instead of collapsing the vertex list;
' a plain Shape keeps the straight-segment GetVertices path bit-for-bit (regression-safe).
' Returns an uninitialised array on failure (caller checks HasAtLeast).
Private Function GetBoundaryVertices(ByVal oRegion As Element, _
                                     ByVal dStrokeTol As Double, _
                                     ByVal dCollinearTol As Double) As Point3d()
    On Error GoTo ErrorHandler
    Dim emptyPts() As Point3d

    ' ComplexShape: densify lines + arcs via GetSubElements.
    If oRegion.IsComplexShapeElement Then
        GetBoundaryVertices = GetComplexBoundaryVertices(oRegion, dStrokeTol, dCollinearTol)
        Exit Function
    End If

    ' Straight-segment Shape: unchanged straight-segment path.
    If Not oRegion.IsVertexList Then
        GetBoundaryVertices = emptyPts
        Exit Function
    End If
    GetBoundaryVertices = oRegion.AsVertexList.GetVertices
    Exit Function

ErrorHandler:
    GetBoundaryVertices = emptyPts
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.GetBoundaryVertices"
End Function

' ============================================================
'  ARC-AWARE BOUNDARY EXTRACTION
' ============================================================

' GetComplexBoundaryVertices
' Walks a ComplexShape boundary via GetSubElements (each sub-element is a ChainableElement —
' a line or an arc; doc ComplexShapeElement_Object). Each line contributes its two
' endpoints; each arc contributes a dense chord polyline (StrokeArcSubElement) whose density
' derives from dStrokeTol. Runs are concatenated and de-duplicated within
' dCollinearTol (shared sub-element endpoints) into one closed loop, consumed unchanged by
' GetClosestSegmentIndex / SignedAreaXY / GetEntryPoint. A dense chord near an arc
' point is locally tangent, so its perpendicular is the (near-)radial direction for free.
' Returns an uninitialised array if the loop cannot be read (caller aborts).
Private Function GetComplexBoundaryVertices(ByVal oRegion As Element, _
                                            ByVal dStrokeTol As Double, _
                                            ByVal dCollinearTol As Double) As Point3d()
    On Error GoTo ErrorHandler
    Dim emptyPts() As Point3d

    Dim oEnum As ElementEnumerator
    Set oEnum = oRegion.AsComplexShapeElement.GetSubElements
    If oEnum Is Nothing Then
        GetComplexBoundaryVertices = emptyPts
        Exit Function
    End If

    Dim verts()     As Point3d
    Dim nVerts      As Long    ' running count of accepted vertices
    Dim oSub        As Element
    Dim runPts()    As Point3d ' dense points for the current sub-element
    Dim nRun        As Long
    Dim nRunsAdded  As Long    ' how many sub-element runs have been chained in so far
    nVerts = 0
    nRunsAdded = 0

    ' ElementEnumerator starts BEFORE the first element: MoveNext first, then Current
    ' (doc ElementEnumerator_Object, mirrored from Length.bas:194).
    Do While oEnum.MoveNext
        Set oSub = oEnum.Current
        If Not oSub Is Nothing Then
            nRun = 0
            If oSub.IsArcElement Then
                ' Arc side: dense chord polyline (radial perpendicular falls out).
                runPts = StrokeArcSubElement(oSub.AsArcElement, dStrokeTol, dCollinearTol, nRun)
            ElseIf oSub.IsVertexList Then
                ' Line OR LineString (broken line): take ALL vertices, not just the two
                ' endpoints. A multi-segment broken line must keep its interior corners --
                ' reducing it to a single Start->End chord injects a chord that cuts across
                ' the corners, corrupting closest-edge selection and the boolean cut on
                ' broken-line sides (the remaining ComplexShape failure mode).
                runPts = oSub.AsVertexList.GetVertices
                nRun = UBound(runPts) + 1
            ElseIf oSub.IsChainableElement Then
                ' Any other straight chainable span: its two endpoints define the edge.
                ReDim runPts(0 To 1)
                runPts(0) = oSub.AsChainableElement.StartPoint
                runPts(1) = oSub.AsChainableElement.EndPoint
                nRun = 2
            End If
            ' A degenerate sub-element (nRun = 0, e.g. zero-sweep arc) is skipped;
            ' if every sub-element is skipped the loop stays empty and the caller aborts.
            If nRun >= 2 Then
                ' Sub-elements share endpoints but are NOT guaranteed to be stored in a
                ' consistent head-to-tail direction; orient each run so it chains continuously
                ' onto the boundary built so far. A reversed sub-element would otherwise inject
                ' a parasitic jump segment that corrupts closest-edge selection, the interior
                ' direction, and the boolean cut (the ComplexShape failure modes).
                OrientRunForChaining verts, nVerts, runPts, nRun, nRunsAdded, dCollinearTol
                AppendBoundaryPoints verts, nVerts, runPts, nRun, dCollinearTol
                nRunsAdded = nRunsAdded + 1
            End If
        End If
    Loop

    If nVerts < 2 Then
        GetComplexBoundaryVertices = emptyPts
        Exit Function
    End If

    ' Force-close the loop if the last accepted vertex does not already coincide with the
    ' first: the boolean still needs a closed boundary, and SignedAreaXY wraps anyway.
    If Point3dDistanceXY(verts(0), verts(nVerts - 1)) > dCollinearTol Then
        ReDim Preserve verts(0 To nVerts)
        verts(nVerts) = verts(0)
        nVerts = nVerts + 1
    End If

    ' Trim to the exact populated length (ReDim Preserve grows in AppendBoundaryPoints).
    ReDim Preserve verts(0 To nVerts - 1)
    GetComplexBoundaryVertices = verts
    Exit Function

ErrorHandler:
    GetComplexBoundaryVertices = emptyPts
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.GetComplexBoundaryVertices"
End Function

' StrokeArcSubElement
' Densifies one arc sub-element into a dense chord polyline (pure-geometry stroker, no
' curve-eval API dependency). Emits nChords+1 points on the arc;
' nChords is DERIVED from dStrokeTol (max chordal deviation) and the arc radius/sweep,
' clamped to [ARES_ARC_MIN_CHORDS, ARES_ARC_MAX_CHORDS] — no literal chord count appears.
'
' Each point is built in the arc's LOCAL parametric frame and mapped to world space by the
' arc's Rotation matrix, so the SAME formula handles a circular arc (PrimaryRadius ==
' SecondaryRadius) AND an elliptical arc: local = (PrimaryRadius*cos t,
' SecondaryRadius*sin t, 0); world = CenterPoint + Rotation * local. (For a circular arc the
' two radii are equal, so this reduces to the centre + R*(cos,sin) form.)
'
' Returns an uninitialised array with nOut = 0 for a degenerate arc (radius <= collinear tol
' or |sweep| ~ 0): the caller skips that sub-element. nOut is the populated count.
Private Function StrokeArcSubElement(ByVal oArc As ArcElement, _
                                     ByVal dStrokeTol As Double, _
                                     ByVal dCollinearTol As Double, _
                                     ByRef nOut As Long) As Point3d()
    On Error GoTo ErrorHandler
    Dim emptyPts() As Point3d
    nOut = 0

    Dim rPrim  As Double
    Dim rSec   As Double
    Dim sweep  As Double
    Dim rMajor As Double
    rPrim = oArc.PrimaryRadius
    rSec = oArc.SecondaryRadius
    sweep = oArc.SweepAngle
    rMajor = rPrim
    If rSec > rMajor Then rMajor = rSec   ' chordal deviation is largest on the major radius

    ' Degenerate arc guard: a vanishing radius or sweep cannot be sampled.
    If rMajor <= dCollinearTol Then
        StrokeArcSubElement = emptyPts
        Exit Function
    End If
    If Abs(sweep) < dCollinearTol Then
        StrokeArcSubElement = emptyPts
        Exit Function
    End If

    ' Chord count from the chordal-deviation tolerance: for a circle of radius R, a chord
    ' subtending angle dTheta deviates from the arc by R*(1 - cos(dTheta/2)). Solving for the
    ' max dTheta that keeps that deviation <= dStrokeTol gives
    '   dThetaMax = 2 * acos(1 - dStrokeTol / R)   (R = rMajor, the worst case for an ellipse)
    ' and nChords = Ceil(|sweep| / dThetaMax). acos requires its argument in [-1, 1]: when
    ' dStrokeTol >= R the tolerance is coarser than the arc, so the minimum chord count
    ' already satisfies it.
    Dim nChords As Long
    Dim dRatio  As Double
    Dim dThetaMax As Double
    dRatio = 1# - (dStrokeTol / rMajor)
    If dRatio <= -1# Then
        nChords = ARES_ARC_MIN_CHORDS
    Else
        dThetaMax = 2# * ArcCosSafe(dRatio)
        If dThetaMax <= 0# Then
            nChords = ARES_ARC_MAX_CHORDS
        Else
            nChords = CLng(Int(Abs(sweep) / dThetaMax)) + 1   ' Ceil for a positive quotient
        End If
    End If
    If nChords < ARES_ARC_MIN_CHORDS Then nChords = ARES_ARC_MIN_CHORDS
    If nChords > ARES_ARC_MAX_CHORDS Then nChords = ARES_ARC_MAX_CHORDS

    Dim oCenter As Point3d
    Dim oRot    As Matrix3d
    Dim startA  As Double
    oCenter = oArc.CenterPoint
    oRot = oArc.Rotation
    startA = oArc.StartAngle

    Dim pts() As Point3d
    ReDim pts(0 To nChords)
    Dim k     As Long
    Dim t     As Double
    Dim ptLocal As Point3d
    Dim ptWorld As Point3d
    For k = 0 To nChords
        t = startA + sweep * (CDbl(k) / CDbl(nChords))
        ' Local parametric point on the (possibly elliptical) arc, then rotate into world.
        ptLocal = Point3dFromXY(rPrim * Cos(t), rSec * Sin(t))
        ptWorld = Point3dFromMatrix3dTimesPoint3d(oRot, ptLocal)
        pts(k) = Point3dAdd(oCenter, ptWorld)
    Next k

    nOut = nChords + 1
    StrokeArcSubElement = pts
    Exit Function

ErrorHandler:
    nOut = 0
    StrokeArcSubElement = emptyPts
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.StrokeArcSubElement"
End Function

' ArcCosSafe
' acos via Atn (VBA lacks a native Acos). Clamps the argument to [-1, 1] so floating-point
' drift just outside the domain does not raise an error. acos(x) = Atn(-x / Sqr(1 - x^2)) +
' PI/2 for |x| < 1; the endpoints are returned directly.
Private Function ArcCosSafe(ByVal x As Double) As Double
    On Error GoTo ErrorHandler
    Dim v As Double
    v = x
    If v <= -1# Then
        ArcCosSafe = Application.PI
    ElseIf v >= 1# Then
        ArcCosSafe = 0#
    Else
        ArcCosSafe = Atn(-v / Sqr(1# - v * v)) + (Application.PI / 2#)
    End If
    Exit Function

ErrorHandler:
    ArcCosSafe = 0#
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.ArcCosSafe"
End Function

' AppendBoundaryPoints
' Concatenates one sub-element's dense run (src(0..nSrc-1)) onto the running boundary
' verts(0..nVerts-1), dropping a leading point that coincides with the last accepted vertex
' within dCollinearTol (shared sub-element endpoint) so SignedAreaXY / GetClosestSegmentIndex
' see a clean loop. Grows verts via ReDim Preserve; updates nVerts.
Private Sub AppendBoundaryPoints(ByRef verts() As Point3d, _
                                 ByRef nVerts As Long, _
                                 ByRef src() As Point3d, _
                                 ByVal nSrc As Long, _
                                 ByVal dCollinearTol As Double)
    On Error GoTo ErrorHandler
    If nSrc <= 0 Then Exit Sub

    Dim i     As Long
    Dim bSkip As Boolean
    For i = 0 To nSrc - 1
        ' Drop a point coincident with the previous accepted vertex (shared endpoint or a
        ' zero-length chord), so consecutive duplicates never reach the cut pipeline.
        bSkip = False
        If nVerts > 0 Then
            If Point3dDistanceXY(verts(nVerts - 1), src(i)) <= dCollinearTol Then bSkip = True
        End If
        If Not bSkip Then
            ReDim Preserve verts(0 To nVerts)
            verts(nVerts) = src(i)
            nVerts = nVerts + 1
        End If
    Next i
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.AppendBoundaryPoints"
End Sub

' OrientRunForChaining
' A ComplexShape's sub-elements share endpoints but are NOT guaranteed to be stored in a
' consistent head-to-tail direction. Before a run is appended, flip it (and, for the second
' run, possibly the first run already in verts) so the boundary stays a single continuous loop.
'   - First run (nRunsAdded = 0): kept as-is; it sets the starting direction.
'   - Second run (nRunsAdded = 1): the first run's correct orientation is still unknown, so
'     test all four endpoint pairings; if the nearest join is at the START of verts, reverse
'     verts; if at the END of the run, reverse the run.
'   - Later runs: reverse the run when its END is nearer the running tail than its START.
Private Sub OrientRunForChaining(ByRef verts() As Point3d, _
                                 ByVal nVerts As Long, _
                                 ByRef src() As Point3d, _
                                 ByVal nSrc As Long, _
                                 ByVal nRunsAdded As Long, _
                                 ByVal dCollinearTol As Double)
    On Error GoTo ErrorHandler
    If nRunsAdded = 0 Or nVerts < 1 Or nSrc < 2 Then Exit Sub

    If nRunsAdded = 1 Then
        ' Resolve the join against BOTH ends of the (single) first run.
        Dim dES As Double, dEE As Double, dSS As Double, dSE As Double
        dES = Point3dDistanceXY(verts(nVerts - 1), src(0))          ' tail -> run start (ideal)
        dEE = Point3dDistanceXY(verts(nVerts - 1), src(nSrc - 1))   ' tail -> run end
        dSS = Point3dDistanceXY(verts(0), src(0))                   ' head -> run start
        dSE = Point3dDistanceXY(verts(0), src(nSrc - 1))            ' head -> run end

        Dim dMin As Double
        Dim bRevVerts As Boolean, bRevSrc As Boolean
        dMin = dES : bRevVerts = False : bRevSrc = False
        If dEE < dMin Then dMin = dEE : bRevVerts = False : bRevSrc = True
        If dSS < dMin Then dMin = dSS : bRevVerts = True : bRevSrc = False
        If dSE < dMin Then dMin = dSE : bRevVerts = True : bRevSrc = True

        If bRevVerts Then ReversePointRun verts, nVerts
        If bRevSrc Then ReversePointRun src, nSrc
    Else
        ' Chain onto the tail: reverse the run if its end is the nearer join.
        If Point3dDistanceXY(verts(nVerts - 1), src(nSrc - 1)) < _
           Point3dDistanceXY(verts(nVerts - 1), src(0)) Then
            ReversePointRun src, nSrc
        End If
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.OrientRunForChaining"
End Sub

' ReversePointRun
' Reverses the first n points of pts in place (used to flip a sub-element run, or the first
' accumulated run, so the boundary chains continuously -- see OrientRunForChaining).
Private Sub ReversePointRun(ByRef pts() As Point3d, ByVal n As Long)
    On Error GoTo ErrorHandler
    Dim i As Long, j As Long
    Dim tmp As Point3d
    i = 0
    j = n - 1
    Do While i < j
        tmp = pts(i)
        pts(i) = pts(j)
        pts(j) = tmp
        i = i + 1
        j = j - 1
    Loop
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.ReversePointRun"
End Sub

' ============================================================
'  CUT GEOMETRY
' ============================================================

' GetEntryPoint
' Perpendicular foot of ClickPt on segment segA->segB, clamped to the segment span. The click
' is accepted wherever MicroStation already located the region (its native, zoom-aware locate
' tolerance); no extra master-unit snap gate is applied -- cut precision is not required, so any
' located click simply cuts at the nearest boundary.
' Uses native Point3dProjectToRay3d: with a normalized Ray.Direction, Parameter is the
' signed distance from Ray.Origin (segA) to the foot.
Private Function GetEntryPoint(ByRef segA As Point3d, _
                               ByRef segB As Point3d, _
                               ByRef ClickPt As Point3d, _
                               ByRef outEntry As Point3d) As Boolean
    On Error GoTo ErrorHandler

    GetEntryPoint = False

    Dim segLen As Double
    segLen = Point3dDistanceXY(segA, segB)
    If segLen <= 0 Then Exit Function

    ' Normalized ray along the segment so Parameter == signed distance from segA.
    Dim ray   As Ray3d
    Dim dirSeg As Point3d
    dirSeg = Point3dSubtract(segB, segA)
    ray.Origin = segA
    ray.Direction = Point3dScale(dirSeg, 1# / segLen)

    Dim param As Double
    Dim foot  As Point3d
    foot = Point3dProjectToRay3d(param, ClickPt, ray)

    ' Clamp the foot to the finite segment [segA, segB].
    If param < 0 Then
        foot = segA
    ElseIf param > segLen Then
        foot = segB
    End If

    outEntry = foot
    GetEntryPoint = True
    Exit Function

ErrorHandler:
    GetEntryPoint = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.GetEntryPoint"
End Function

' GetInteriorDirection
' Orients the cut perpendicular to segA->segB into the region interior from
' the boundary ORIENTATION rather than a distance-based sample. Sampling at a fraction of
' the segment length failed on narrow regions: a long edge of a thin strip pushed the test
' point clean through to the far side, so BOTH perpendicular signs read "outside" and the
' cut was wrongly rejected. Perp2D returns the left-hand (90 CCW) perpendicular of
' segA->segB; for a CCW boundary (signed area > 0) the interior lies on the left, so that
' perpendicular already points inward; for a CW boundary it is negated. Deterministic for
' convex and concave simple polygons alike. Returns a unit inward vector, or False when the
' perpendicular or the boundary area is degenerate.
Private Function GetInteriorDirection(ByRef segA As Point3d, _
                                      ByRef segB As Point3d, _
                                      ByRef verts() As Point3d, _
                                      ByRef outDir As Point3d) As Boolean
    On Error GoTo ErrorHandler

    GetInteriorDirection = False

    ' Unit left-hand (90 CCW) perpendicular. Perp2D returns a zero vector for a zero-length
    ' segment (caller already rejected degenerate segments; defence-in-depth here).
    Dim perp As Point3d
    perp = Geometry.Perp2D(segA, segB, 1#)
    If Point3dMagnitudeSquared(perp) < 1E-24 Then Exit Function

    ' Boundary orientation picks the interior side: CCW (> 0) => interior on the left of each
    ' directed edge => left-hand perpendicular is inward; CW (< 0) => negate. A near-zero
    ' area is a degenerate boundary that cannot be split (leave the result False).
    Dim dArea As Double
    dArea = SignedAreaXY(verts)
    If dArea > 0# Then
        outDir = perp
        GetInteriorDirection = True
    ElseIf dArea < 0# Then
        outDir = Point3dNegate(perp)
        GetInteriorDirection = True
    End If
    Exit Function

ErrorHandler:
    GetInteriorDirection = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.GetInteriorDirection"
End Function

' RefineDirectionRadialIfArc
' When the cut entry lies on an arc side of a ComplexShape this does two things, so the cut
' follows the arc's angle toward its origin (user preference) AND severs a concave (interior)
' arc cleanly:
'   1. SNAP entryPt from the inset stroked chord back onto the REAL arc (radial projection onto
'      the possibly-elliptical arc). The chord sits up to dStrokeTol toward the centre; on an
'      INTERIOR/concave arc that puts entryPt in the hole, OUTSIDE the region, so the exit ray
'      would re-cross the very same arc a hair away and build a zero-width knife -> "fewer than
'      two regions (1)". Snapping onto the true boundary lets the exit ray clear the start arc
'      and reach the opposite side. (Circular arcs reduce to centre + r * unit(entry - centre).)
'   2. Set dirIn to the EXACT radial axis (entry -> arc centre) instead of the chord normal,
'      keeping the inward SENSE of the perpendicular direction -- the two are near-collinear so
'      that dot-product sign test is well-conditioned.
' No-op for a plain Shape, when the entry is not on any arc sub-element (a straight side), or at
' the arc centre. Failure is non-fatal: entryPt/dirIn are left as computed by the chord path.
Private Sub RefineDirectionRadialIfArc(ByVal oRegion As Element, _
                                       ByRef entryPt As Point3d, _
                                       ByVal dStrokeTol As Double, _
                                       ByVal dCollinearTol As Double, _
                                       ByRef dirIn As Point3d)
    On Error GoTo ErrorHandler
    If Not oRegion.IsComplexShapeElement Then Exit Sub

    Dim oArc As ArcElement
    Set oArc = GetArcAtPoint(oRegion, entryPt, dStrokeTol, dCollinearTol)
    If oArc Is Nothing Then Exit Sub

    Dim center As Point3d
    center = oArc.CenterPoint

    ' --- (1) Snap entryPt radially onto the real arc (handles circular AND elliptical). ---
    Dim rPrim As Double, rSec As Double
    rPrim = oArc.PrimaryRadius
    rSec = oArc.SecondaryRadius
    If rPrim > dCollinearTol And rSec > dCollinearTol Then
        Dim oRot   As Matrix3d
        Dim vMinus As Point3d
        Dim locP   As Point3d
        Dim denom  As Double
        oRot = oArc.Rotation
        vMinus = Point3dSubtract(entryPt, center)
        ' Into the arc's local frame (Rotation is orthonormal: transpose == inverse).
        locP = Point3dFromMatrix3dTransposeTimesPoint3d(oRot, vMinus)
        denom = (locP.X / rPrim) * (locP.X / rPrim) + (locP.Y / rSec) * (locP.Y / rSec)
        If denom > 0# Then
            Dim s     As Double
            Dim locOn As Point3d
            s = 1# / Sqr(denom)                       ' scale that lands (locP) on the ellipse
            locOn = Point3dFromXY(s * locP.X, s * locP.Y)
            entryPt = Point3dAdd(center, Point3dFromMatrix3dTimesPoint3d(oRot, locOn))
        End If
    End If

    ' --- (2) Radial cut axis through the (snapped) entry and the arc centre, kept inward. ---
    Dim rx As Double, ry As Double, rlen As Double
    rx = center.X - entryPt.X
    ry = center.Y - entryPt.Y
    rlen = Sqr(rx * rx + ry * ry)
    If rlen <= dCollinearTol Then Exit Sub   ' entry at the centre: keep the perpendicular
    rx = rx / rlen
    ry = ry / rlen
    If (rx * dirIn.X + ry * dirIn.Y) < 0# Then
        rx = -rx
        ry = -ry
    End If
    dirIn = Point3dFromXY(rx, ry)
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.RefineDirectionRadialIfArc"
End Sub

' GetArcAtPoint
' Scans the ComplexShape's arc sub-elements; if pt lies on one (within dStrokeTol of that arc's
' own stroked chord polyline -- pt itself came from the global stroking, so the match distance
' is ~0), returns that ArcElement. The first matching arc wins (a shared arc/line corner is an
' acceptable tie). Nothing when pt sits on a straight side (no arc match).
Private Function GetArcAtPoint(ByVal oRegion As Element, _
                               ByRef pt As Point3d, _
                               ByVal dStrokeTol As Double, _
                               ByVal dCollinearTol As Double) As ArcElement
    On Error GoTo ErrorHandler
    Set GetArcAtPoint = Nothing

    Dim oEnum As ElementEnumerator
    Set oEnum = oRegion.AsComplexShapeElement.GetSubElements
    If oEnum Is Nothing Then Exit Function

    Dim oSub     As Element
    Dim runPts() As Point3d
    Dim nRun     As Long
    Dim i        As Long
    Do While oEnum.MoveNext
        Set oSub = oEnum.Current
        If Not oSub Is Nothing Then
            If oSub.IsArcElement Then
                nRun = 0
                runPts = StrokeArcSubElement(oSub.AsArcElement, dStrokeTol, dCollinearTol, nRun)
                For i = 0 To nRun - 2
                    If DistPointToSegmentXY(pt, runPts(i), runPts(i + 1)) <= dStrokeTol Then
                        Set GetArcAtPoint = oSub.AsArcElement
                        Exit Function
                    End If
                Next i
            End If
        End If
    Loop
    Exit Function

ErrorHandler:
    Set GetArcAtPoint = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.GetArcAtPoint"
End Function

' GetExitPoint
' Casts a long ray from entryPt along dirIn and intersects it with the region boundary
' (AsIntersectableElement.GetIntersectionPoints, the Length.bas ray-cast idiom). Returns
' the crossing with the smallest positive distance strictly beyond dCollinearTol from the
' entry. False if there is no such crossing (clean abort).
Private Function GetExitPoint(ByVal oRegion As Element, _
                              ByRef entryPt As Point3d, _
                              ByRef dirIn As Point3d, _
                              ByVal dCollinearTol As Double, _
                              ByRef outExit As Point3d) As Boolean
    On Error GoTo ErrorHandler

    GetExitPoint = False
    If Not oRegion.IsIntersectableElement Then Exit Function

    ' Ray length: comfortably spans the whole region (bbox diagonal + margin).
    Dim oRange  As Range3d
    Dim dRayLen As Double
    oRange = oRegion.Range
    dRayLen = Point3dDistanceXY(oRange.Low, oRange.High) + 1#

    Dim farPt As Point3d
    farPt = Point3dAddScaled(entryPt, dirIn, dRayLen)

    Dim oRayEl As LineElement
    Set oRayEl = CreateLineElement2(Nothing, entryPt, farPt)

    Dim oIsect As IntersectableElement
    Set oIsect = oRegion.AsIntersectableElement

    Dim isectPts() As Point3d
    Dim nIsect     As Long
    On Error Resume Next
    isectPts = oIsect.GetIntersectionPoints(oRayEl, Matrix3dIdentity)
    nIsect = -1
    nIsect = UBound(isectPts)
    On Error GoTo ErrorHandler

    If nIsect < 0 Then Exit Function

    ' Pick the nearest crossing whose along-ray distance is strictly beyond the entry.
    Dim i        As Long
    Dim dAlong   As Double
    Dim dBest    As Double
    Dim bFound   As Boolean
    bFound = False
    dBest = 0
    For i = 0 To nIsect
        dAlong = AlongRayDistance(entryPt, dirIn, isectPts(i))
        If dAlong > dCollinearTol Then
            If Not bFound Or dAlong < dBest Then
                dBest = dAlong
                outExit = isectPts(i)
                bFound = True
            End If
        End If
    Next i

    GetExitPoint = bFound
    Exit Function

ErrorHandler:
    GetExitPoint = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.GetExitPoint"
End Function

' AlongRayDistance
' Signed distance of pt projected onto the ray (origin, unit dir). dir is expected
' unit-length (GetInteriorDirection returns a unit vector).
Private Function AlongRayDistance(ByRef origin As Point3d, _
                                  ByRef dir As Point3d, _
                                  ByRef pt As Point3d) As Double
    On Error GoTo ErrorHandler
    Dim v As Point3d
    v = Point3dSubtract(pt, origin)
    AlongRayDistance = v.X * dir.X + v.Y * dir.Y
    Exit Function

ErrorHandler:
    AlongRayDistance = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.AlongRayDistance"
End Function

' ============================================================
'  KNIFE + BOOLEAN
' ============================================================

' BuildKnife
' Builds a thin closed rectangle (ShapeElement) straddling the chord entryPt -> exitPt:
' offset the chord by a tiny half-width (ARES_KNIFE_HALFWIDTH_FACTOR * dCollinearTol, NOT a
' literal) on each side, and over-extend slightly past each end so the difference fully
' severs the region. Reuses the 4-point rectangle construction from BuildLineZone's
' flat-cap branch (Zoning.bas:794-796).
Private Function BuildKnife(ByRef entryPt As Point3d, _
                            ByRef exitPt As Point3d, _
                            ByVal dCollinearTol As Double, _
                            ByVal dStrokeTol As Double, _
                            ByVal dRegionDiag As Double) As Element
    On Error GoTo ErrorHandler

    Set BuildKnife = Nothing

    Dim chordLen As Double
    chordLen = Point3dDistanceXY(entryPt, exitPt)
    If chordLen <= dCollinearTol Then Exit Function

    ' Half-width: the larger of the absolute floor (Collinear_Tol * factor, for tiny regions
    ' near the origin) and an extent-proportional term (bbox diagonal * rel factor). The latter
    ' keeps the slot above GetRegionDifference's extent-scaled cleanup tolerance on large
    ' regions, where an absolute-only width collapses to a single region.
    Dim hw     As Double
    Dim hwRel  As Double
    Dim perp   As Point3d   ' length = half-width, perpendicular to the chord
    hw = dCollinearTol * ARES_KNIFE_HALFWIDTH_FACTOR
    hwRel = dRegionDiag * ARES_KNIFE_HALFWIDTH_REL_FACTOR
    If hwRel > hw Then hw = hwRel
    perp = Geometry.Perp2D(entryPt, exitPt, hw)
    If Point3dMagnitudeSquared(perp) < 1E-24 Then Exit Function

    ' Over-extend each end along the chord so the slot crosses the boundary cleanly. On an arc
    ' side the entry point sits up to dStrokeTol inside the true curved boundary (it is the foot
    ' on the stroked chord, not on the real arc), so the chord-proportional over-extension is
    ' floored at a small multiple of dStrokeTol AND at the same extent-proportional term as the
    ' half-width; otherwise a narrow strip on a large region leaves an uncut bridge and the
    ' difference returns a single region.
    Dim over    As Double
    Dim overMin As Double
    Dim dirAxis As Point3d
    over = chordLen * ARES_KNIFE_OVEREXTEND_FACTOR
    overMin = dStrokeTol * ARES_KNIFE_ARC_OVEREXTEND_FACTOR
    If dRegionDiag * ARES_KNIFE_HALFWIDTH_REL_FACTOR > overMin Then overMin = dRegionDiag * ARES_KNIFE_HALFWIDTH_REL_FACTOR
    If over < overMin Then over = overMin
    dirAxis = Point3dScale(Point3dSubtract(exitPt, entryPt), 1# / chordLen)

    Dim p0 As Point3d   ' extended entry end
    Dim p1 As Point3d   ' extended exit end
    p0 = Point3dAddScaled(entryPt, dirAxis, -over)
    p1 = Point3dAddScaled(exitPt, dirAxis, over)

    ' Four rectangle corners (left side then right side), closed back to the first.
    Dim L0 As Point3d, L1 As Point3d, R1 As Point3d, R0 As Point3d
    L0 = Point3dAdd(p0, perp)
    L1 = Point3dAdd(p1, perp)
    R1 = Point3dSubtract(p1, perp)
    R0 = Point3dSubtract(p0, perp)

    Dim rectPts(0 To 4) As Point3d
    rectPts(0) = L0 : rectPts(1) = L1 : rectPts(2) = R1 : rectPts(3) = R0 : rectPts(4) = L0
    Set BuildKnife = CreateShapeElement1(Nothing, rectPts)
    Exit Function

ErrorHandler:
    Set BuildKnife = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.BuildKnife"
End Function

' SplitByKnife
' Performs region = region - knife via GetRegionDifference, evaluated near the origin
' (Zoning precision workaround): clone the region, translate clone + knife by
' -region.Range.High, run the boolean, then translate each result back by +High.
' Accumulates the resulting halves into outHalves(); returns False on a hard failure.
Private Function SplitByKnife(ByVal oRegion As Element, _
                              ByVal knifeEl As Element, _
                              ByRef outHalves() As Element, _
                              ByRef nHalves As Long) As Boolean
    On Error GoTo ErrorHandler

    SplitByKnife = False
    nHalves = 0

    ' Translate both operands near the origin (GetRegionDifference is unreliable at large
    ' DGN coordinates — MicroStation bug). Mirror Zoning.bas:187-217 exactly.
    Dim toOrigin   As Point3d
    Dim fromOrigin As Point3d
    toOrigin = Point3dNegate(oRegion.Range.High)
    fromOrigin = Point3dNegate(toOrigin)

    Dim regionClone As Element
    Set regionClone = oRegion.Clone
    regionClone.Move toOrigin
    knifeEl.Move toOrigin

    Dim solid(0 To 0) As Element
    Dim holes(0 To 0) As Element
    Set solid(0) = regionClone
    Set holes(0) = knifeEl

    Dim oEnum As ElementEnumerator
    Set oEnum = GetRegionDifference(solid, holes, Nothing, msdFillModeNotFilled)
    If oEnum Is Nothing Then Exit Function

    Dim oHalf As Element
    Do While oEnum.MoveNext
        Set oHalf = oEnum.Current
        If Not oHalf Is Nothing Then
            oHalf.Move fromOrigin
            ReDim Preserve outHalves(0 To nHalves)
            Set outHalves(nHalves) = oHalf
            nHalves = nHalves + 1
        End If
    Loop

    SplitByKnife = True
    Exit Function

ErrorHandler:
    SplitByKnife = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.SplitByKnife"
End Function

' ============================================================
'  OUTPUT
' ============================================================

' WriteHalves
' Applies the original's level + symbology (read from oRegion) to every half and adds it
' to the active model. Called only after >= 2 halves are validated; the original is
' deleted by the caller AFTER this returns True (add-both-then-delete ordering).
'
' This is a Function returning True ONLY if BOTH halves were added and styled
' successfully. Any AddElement / property-set / Rewrite failure routes to ErrorHandler and
' returns False, so the caller does NOT delete the original (anti-destructive on the error
' path too). Edge case (recorded in Dev Notes): if the first half is added then the second
' fails, the first half is already in the model and — because this returns False — the
' original is KEPT, so the user sees one extra region overlapping the original (non-
' destructive and visible) rather than a silent data loss. A transactional rollback of the
' first AddElement is out of scope (no undo-grouping work; MicroStation's native command
' transaction covers interactive undo).
'
' Since MicroStation 8.1 Level and LineStyle are object-valued (read into locals with Set;
' written onto the element by reference like Zoning.ApplySym, no Set). The Level property
' can only be set once the element is a model member, so each half is AddElement'd FIRST,
' then its level + symbology applied.
Private Function WriteHalves(ByVal oRegion As Element, _
                             ByRef halves() As Element, _
                             ByVal nHalves As Long) As Boolean
    On Error GoTo ErrorHandler
    WriteHalves = False

    Dim srcLevel  As Level
    Dim srcStyle  As LineStyle
    Dim srcColor  As Long
    Dim srcWeight As Long
    Set srcLevel = oRegion.Level
    Set srcStyle = oRegion.LineStyle
    srcColor = oRegion.Color
    srcWeight = oRegion.LineWeight

    Dim i        As Long
    Dim nWritten As Long
    nWritten = 0
    For i = 0 To nHalves - 1
        If Not halves(i) Is Nothing Then
            ' Add first: Level cannot be assigned to a non-member element (see doc).
            ActiveModelReference.AddElement halves(i)
            halves(i).Level = srcLevel
            halves(i).Color = srcColor
            halves(i).LineStyle = srcStyle
            halves(i).LineWeight = srcWeight
            halves(i).Rewrite
            nWritten = nWritten + 1
        End If
    Next i

    ' True only if BOTH halves really made it into the model.
    WriteHalves = (nWritten >= 2)
    Exit Function

ErrorHandler:
    WriteHalves = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.WriteHalves"
End Function

' ============================================================
'  LOCAL GEOMETRY HELPERS
' ============================================================

' GetClosestSegmentIndex
' Index s (into verts) of the boundary segment verts(s)->verts(s+1) whose XY distance to
' ClickPt is smallest. Computed locally from the already-resolved vertex array instead of
' the native VertexList.GetClosestSegment, which raised a COM collection error on some
' region elements. Returns -1 when there are fewer than two vertices.
Private Function GetClosestSegmentIndex(ByRef verts() As Point3d, ByRef ClickPt As Point3d) As Long
    On Error GoTo ErrorHandler
    Dim i     As Long
    Dim dDist As Double
    Dim dBest As Double
    Dim nBest As Long
    nBest = -1
    For i = LBound(verts) To UBound(verts) - 1
        dDist = DistPointToSegmentXY(ClickPt, verts(i), verts(i + 1))
        If nBest < 0 Or dDist < dBest Then
            dBest = dDist
            nBest = i
        End If
    Next i
    GetClosestSegmentIndex = nBest
    Exit Function

ErrorHandler:
    GetClosestSegmentIndex = -1
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.GetClosestSegmentIndex"
End Function

' DistPointToSegmentXY
' Shortest XY distance from point P to the finite segment A->B, with the projection
' parameter clamped to [0,1] so the result is the distance to the segment (not the infinite
' line). A zero-length segment falls back to the distance to A.
Private Function DistPointToSegmentXY(ByRef P As Point3d, ByRef A As Point3d, ByRef B As Point3d) As Double
    On Error GoTo ErrorHandler
    Dim abx As Double, aby As Double
    Dim apx As Double, apy As Double
    Dim L2  As Double, t   As Double
    Dim cx  As Double, cy  As Double
    abx = B.X - A.X
    aby = B.Y - A.Y
    apx = P.X - A.X
    apy = P.Y - A.Y
    L2 = abx * abx + aby * aby
    If L2 <= 1E-24 Then
        DistPointToSegmentXY = Sqr(apx * apx + apy * apy)
        Exit Function
    End If
    t = (apx * abx + apy * aby) / L2
    If t < 0# Then
        t = 0#
    ElseIf t > 1# Then
        t = 1#
    End If
    cx = A.X + t * abx
    cy = A.Y + t * aby
    DistPointToSegmentXY = Sqr((P.X - cx) * (P.X - cx) + (P.Y - cy) * (P.Y - cy))
    Exit Function

ErrorHandler:
    DistPointToSegmentXY = 1E+30
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.DistPointToSegmentXY"
End Function

' SignedAreaXY
' Signed XY area of the closed polygon described by verts (shoelace). Positive for a
' counter-clockwise vertex order, negative for clockwise. The loop is treated as closed
' (last vertex wraps to the first), so a duplicated closing vertex contributes zero and is
' harmless. Used by GetInteriorDirection to choose the interior side of a boundary segment.
Private Function SignedAreaXY(ByRef verts() As Point3d) As Double
    On Error GoTo ErrorHandler
    Dim i    As Long
    Dim j    As Long
    Dim lo   As Long
    Dim hi   As Long
    Dim dSum As Double
    lo = LBound(verts)
    hi = UBound(verts)
    dSum = 0#
    For i = lo To hi
        If i = hi Then
            j = lo
        Else
            j = i + 1
        End If
        dSum = dSum + (verts(i).X * verts(j).Y - verts(j).X * verts(i).Y)
    Next i
    SignedAreaXY = dSum * 0.5
    Exit Function

ErrorHandler:
    SignedAreaXY = 0#
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "RegionSplit.SignedAreaXY"
End Function

' HasAtLeast
' Safe lower-bound check: True if arr is initialised and has at least n items.
Private Function HasAtLeast(ByRef arr() As Point3d, ByVal n As Long) As Boolean
    On Error Resume Next
    HasAtLeast = False
    If (UBound(arr) - LBound(arr) + 1) >= n Then HasAtLeast = True
    On Error GoTo 0
End Function

' ShowSplitStatus
' Shows a user-facing status via LangManager.GetTranslation when initialised, else falls
' back to a literal English string (as several Command.bas subs do).
Private Sub ShowSplitStatus(ByVal Key As String, ByVal FallbackEN As String)
    On Error Resume Next
    If LangManager.IsInit Then
        ShowStatus GetTranslation(Key)
    Else
        ShowStatus FallbackEN
    End If
    On Error GoTo 0
End Sub
