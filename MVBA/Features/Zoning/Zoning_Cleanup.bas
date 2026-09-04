' Module: Zoning_Cleanup
' Description: Contour cleanup for the Zoning split - the passes that run on what is actually
'              drawn, plus the cell walk for a zone that has a hole. Everything here is
'              cosmetic and must never cost a zone: each pass hands its input back
'              untouched when it cannot help or anything goes wrong.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ErrorHandler

Option Explicit


' The cleanup factor for the run in progress, read once from config by Zoning and used by the
' cleanup passes, which are reached through WriteEl and have no parameter to carry it. Zero until a
' run sets it, which is the safe default: no run, no cleanup.
Private mdCleanupFactor As Double

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

' The four thresholds below are the cleanup's own settings, and each is multiplied by
' ARES_Zoning_Cleanup_Factor (default 1.0) before use - see mdCleanupFactor. One knob instead of four
' numbers: 0 turns the cleanup off without recompiling, 1 is what was measured on Asketyll's corpus,
' above 1 is more aggressive. The values stay written here as themselves, so each still reads as
' "2 cm at a 2 m offset" instead of becoming meaningless in isolation.
'
' Three things are deliberately NOT on that knob, and must not be:
'   - CAP_OVERLAP_RATIO (in Zoning_Builders) is a numerical necessity, not a matter of taste. It is
'     what stops the union
'     losing whole buffers - proven twice, on lines and on arcs - and a factor of 0 would quietly
'     take zones with it.
'   - SLIVER_SWEEPS is a safety cap on a loop, not a threshold.
'   - COVERAGE_* is the instrument that checks the result. An instrument you can tune until it
'     agrees with you has stopped being one.
'
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

' SetCleanupFactor
' ---------------------------------------------------------------------------
' Hands the cleanup its scale for the run about to start. Public because the value is read from
' config by Zoning, which owns the run, while everything that uses it lives here. A negative factor
' is clamped rather than refused: it can only have come from a typo in a config variable, and no
' cleanup is a better answer to that than a negative threshold.
' ---------------------------------------------------------------------------
Public Sub SetCleanupFactor(ByVal dFactor As Double)
    If dFactor < 0 Then dFactor = 0
    mdCleanupFactor = dFactor
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
Public Function UnwrapLoneCell(ByVal oCell As Element) As Element
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
Public Function CleanContour(ByVal oEl As Element, ByVal Dist As Double) As Element
    Dim oOut As Element
    Dim oWas As Element
    Dim k    As Long

    Set oOut = oEl
    If mdCleanupFactor <= 0 Then
        Set CleanContour = oOut          ' factor 0: the cleanup is off, hand the contour back as it is
        Exit Function
    End If

    If ENABLE_VERTEX_THINNING Then _
        Set oOut = CleanTinyVertices(oOut, Dist * VERTEX_MERGE_RATIO * mdCleanupFactor)

    ' Swept until it settles - see SLIVER_SWEEPS. Dropping a sliver frees its neighbours to become
    ' droppable in turn, and one sweep can only ever see the chain it started with.
    If ENABLE_FLAT_ARC_DROP Then
        For k = 1 To SLIVER_SWEEPS
            Set oWas = oOut
            Set oOut = DropSliverEdges(oOut, FLAT_ARC_DEG * mdCleanupFactor, _
                                       Dist * FLAT_ARC_LEN_RATIO * mdCleanupFactor, _
                                       Dist * VERTEX_MERGE_RATIO * mdCleanupFactor)
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
Public Sub CleanCellChildren(ByVal oCell As CellElement, ByVal Dist As Double)
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

    dMinArea = Dist * Dist * MIN_CELL_PART_AREA_RATIO * mdCleanupFactor
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
Public Function AreaOf(ByVal oEl As Element) As Double
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
