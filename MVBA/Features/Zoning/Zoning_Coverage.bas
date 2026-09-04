' Module: Zoning_Coverage
' Description: The final coverage check for the Zoning split - measures every source element
'              against the zones that were actually produced and rebuilds the pieces that
'              ended up outside them. The only check that looks at the finished product.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: Length, Zoning (DbgLine), Zoning_Builders, Zoning_Cleanup (AreaOf), Zoning_Dispatchers

Option Explicit


' Final coverage check - see RepairUncovered.
'
' A source element counts as uncovered when more of it than this SHARE of the offset distance lies
' outside the merged zones. 0.05 is 10 cm at the 2 m zoning distance: below that we are looking at
' the measurement's own noise - the containment test rides on ray casts and, on arcs, on chord
' midpoints - and not at a hole anyone will ever see.
Private Const COVERAGE_SLACK_RATIO As Double = 0.05

' And the check refuses to act at all when it accuses more than this share of what it measured. A
' zoning run that left half its cables outside their own zones is not a zoning failure, it is a
' broken instrument - that mistake has been made on this module before, three times in a row, and
' each guard built on it made things worse. Refusing costs a run; acting on it corrupts every zone.
Private Const COVERAGE_SANITY_SHARE As Double = 0.5

' ...and only once it has measured at least this many elements. Below that, "most of them" is not a
' statement about anything: one isolated cable that really is uncovered is 100% of the sample, and
' the guard would refuse the very repair it was asked for. Found the hard way - Asketyll isolated a
' single cable in a DGN to test the pass, and the protection made the test impossible.
Private Const COVERAGE_SANITY_MIN As Long = 5

' RepairUncovered
' ---------------------------------------------------------------------------
' Measures every source element against the zones that were actually produced, and rebuilds the
' buffer of each one that is not fully inside them. The rebuilt buffers come back in repBufs()/nRep
' for the caller to merge; nothing is written here.
'
' Why it exists: the union is the one step that can silently lose coverage. A buffer that fails to
' merge does not raise anything, it is simply absent from the result, and a cable then runs outside
' its own zone with no trace anywhere. This is the only check that looks at the finished product.
'
' The instrument is Length.GetPartialLengthInsideZones, the same one that made the Cable Report agree
' with MicroStation's own measurement to the millimetre on a 487 m cable. It is used through its
' public face, not reimplemented: a second containment test would be a second thing to be wrong.
'
' Two things it deliberately does NOT do:
'   - it skips what that instrument cannot measure, cells and ellipses among them. An element whose
'     length cannot be read is left alone, never treated as uncovered: an unmeasurable element and an
'     uncovered one are not the same claim, and confusing them would rebuild the entire drawing.
'   - it refuses to act when it accuses more than COVERAGE_SANITY_SHARE of what it measured, and says
'     so. A check that condemns most of what it looks at is reporting on itself.
' ---------------------------------------------------------------------------
Public Sub RepairUncovered(ByRef Elements() As Element, _
                            ByRef zones() As Element, _
                            ByVal nZones As Long, _
                            ByVal Dist As Double, _
                            ByVal TargetLevel As Level, _
                            ByVal Color As Long, _
                            ByVal Style As String, _
                            ByVal Weight As Long, _
                            ByVal DebugMode As Boolean, _
                            ByVal RoundCaps As Boolean, _
                            ByRef repBufs() As Element, _
                            ByRef nRep As Long)
    On Error GoTo ErrorHandler

    nRep = 0
    If nZones <= 0 Then Exit Sub

    Dim i        As Long
    Dim nTested  As Long
    Dim nMissing As Long
    Dim dTotal   As Double
    Dim dIn      As Double
    Dim dSlack   As Double
    Dim bMiss()  As Boolean

    dSlack = Dist * COVERAGE_SLACK_RATIO
    ReDim bMiss(LBound(Elements) To UBound(Elements))

    For i = LBound(Elements) To UBound(Elements)
        If Measurable(Elements(i)) Then
            dTotal = Length.GetLength(Elements(i), RndLength:=False)
            If dTotal > dSlack Then
                nTested = nTested + 1
                dIn = Length.GetPartialLengthInsideZones(Elements(i), zones)
                If dIn < dTotal - dSlack Then
                    bMiss(i) = True
                    nMissing = nMissing + 1
                    If DebugMode Then _
                        DbgLine "COVER #" & i & " type " & Elements(i).Type & " : " & Format(dIn, "0.000") & _
                                " m inside of " & Format(dTotal, "0.000") & " m -> NOT COVERED"
                ElseIf DebugMode Then
                    DbgLine "COVER #" & i & " type " & Elements(i).Type & " : " & Format(dIn, "0.000") & _
                            " m inside of " & Format(dTotal, "0.000") & " m -> covered"
                End If
            ElseIf DebugMode Then
                DbgLine "COVER #" & i & " type " & Elements(i).Type & " : length " & Format(dTotal, "0.000") & _
                        " m is under the slack, skipped"
            End If
        ElseIf DebugMode Then
            DbgLine "COVER #" & i & " type " & Elements(i).Type & " : not measurable, skipped"
        End If
    Next i

    If DebugMode Then _
        DbgLine "COVER verdict: " & nMissing & " of " & nTested & " measurable element(s) short, against " & _
                nZones & " zone(s)"
    If nMissing = 0 Then Exit Sub

    ' The share test needs a sample big enough for a share to mean something - see COVERAGE_SANITY_MIN.
    If nTested >= COVERAGE_SANITY_MIN And nMissing > nTested * COVERAGE_SANITY_SHARE Then
        ErrorHandler.HandleError "coverage check REFUSED - " & nMissing & " of " & nTested & _
                                 " elements reported outside their own zones. A result like that is " & _
                                 "the measurement failing, not the zoning; the zones are left as they are.", _
                                 0, "", "Zoning.RepairUncovered"
        Exit Sub
    End If

    ' Rebuilt PIECE BY PIECE, never through the dispatcher. Going back through the dispatcher
    ' reproduces the defect exactly: a complex chain fuses its own sub-buffers before emitting, so a
    ' sub-element already lost there is lost again, and the repair hands back the very shapes that
    ' were missing the piece. Measured: a 323 m chain came back as the same two zones it went in as.
    For i = LBound(Elements) To UBound(Elements)
        If bMiss(i) Then
            CollectUncoveredPieces Elements(i), zones, Dist, RoundCaps, dSlack, _
                                   repBufs, nRep, DebugMode
        End If
    Next i

    If DebugMode Then DbgLine "COVER rebuilt " & nRep & " buffer(s) to merge back in"
    Exit Sub

ErrorHandler:
    ' A failure here must not cost the zones that were already merged: hand back nothing to add.
    nRep = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.RepairUncovered"
End Sub

' CollectUncoveredPieces
' ---------------------------------------------------------------------------
' Walks ONE source element down to the pieces its buffers are actually built from - a Line, an Arc,
' one segment of a linestring - measures each against the zones, and appends a buffer for every piece
' that is not inside them. Nothing is written; the caller merges what comes back.
'
' Asketyll's rule, and the reason it has to be the pieces and not the element: the arc in the middle
' of a 323 m chain was the thing missing from the zones, and rebuilding the chain around it rebuilt
' the same hole. Only the piece itself can be put back.
'
' Caps follow the SAME rule as the first pass - CapRoundAt against the chain's own free ends - rather
' than being rounded for safety. A round cap at a genuine free end would push the zone out by the
' whole offset distance, which on RunOutline's 0.2 m is a fifth of the zone and plainly visible.
' ---------------------------------------------------------------------------
Private Sub CollectUncoveredPieces(ByVal oEl As Element, _
                                   ByRef zones() As Element, _
                                   ByVal Dist As Double, _
                                   ByVal RoundCaps As Boolean, _
                                   ByVal dSlack As Double, _
                                   ByRef repBufs() As Element, _
                                   ByRef nRep As Long, _
                                   ByVal DebugMode As Boolean)
    On Error GoTo ErrorHandler

    Dim gStart   As Point3d
    Dim gEnd     As Point3d
    Dim allRound As Boolean
    Dim tol      As Double

    tol = Dist * ARES_CAP_MATCH_FRAC

    ' Same free-end reading as the builders: a closed chain has none, so every cap is rounded.
    If oEl.Type = msdElementTypeComplexShape Or oEl.Type = msdElementTypeShape Then
        allRound = True
    Else
        gStart = oEl.AsChainableElement.StartPoint
        gEnd = oEl.AsChainableElement.EndPoint
        allRound = RoundCaps Or Point3dEqualTolerance(gStart, gEnd, tol)
    End If

    WalkPieces oEl, zones, Dist, dSlack, gStart, gEnd, allRound, tol, repBufs, nRep, DebugMode
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.CollectUncoveredPieces"
End Sub

' WalkPieces
' ---------------------------------------------------------------------------
' The recursive half of CollectUncoveredPieces. Descends a chain into its sub-elements, a linestring
' into its segments, and tests the leaves. The chain's global ends travel down unchanged so a cap is
' decided on the WHOLE cable's geometry, never on the fragment's.
' ---------------------------------------------------------------------------
Private Sub WalkPieces(ByVal oEl As Element, _
                       ByRef zones() As Element, _
                       ByVal Dist As Double, _
                       ByVal dSlack As Double, _
                       ByRef gStart As Point3d, _
                       ByRef gEnd As Point3d, _
                       ByVal allRound As Boolean, _
                       ByVal tol As Double, _
                       ByRef repBufs() As Element, _
                       ByRef nRep As Long, _
                       ByVal DebugMode As Boolean)
    On Error GoTo ErrorHandler

    Dim oEE   As ElementEnumerator
    Dim oVL   As VertexList
    Dim verts() As Point3d
    Dim j     As Long
    Dim oSeg  As Element

    Select Case oEl.Type

        Case msdElementTypeComplexString, msdElementTypeComplexShape
            Set oEE = oEl.AsComplexElement.GetSubElements
            Do While oEE.MoveNext
                WalkPieces oEE.Current, zones, Dist, dSlack, gStart, gEnd, allRound, tol, _
                           repBufs, nRep, DebugMode
            Loop

        Case msdElementTypeLineString, msdElementTypeShape
            Set oVL = oEl
            verts = oVL.GetVertices
            For j = LBound(verts) To UBound(verts) - 1
                Set oSeg = CreateLineElement2(Nothing, verts(j), verts(j + 1))
                TestAndBuffer oSeg, zones, Dist, dSlack, gStart, gEnd, allRound, tol, _
                              repBufs, nRep, DebugMode
            Next j

        Case msdElementTypeLine, msdElementTypeArc
            TestAndBuffer oEl, zones, Dist, dSlack, gStart, gEnd, allRound, tol, _
                          repBufs, nRep, DebugMode

    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.WalkPieces"
End Sub

' TestAndBuffer
' ---------------------------------------------------------------------------
' One leaf piece: measure how much of it lies inside the zones and, if it is short, build its buffer
' with the builder the first pass would have used and append it.
' ---------------------------------------------------------------------------
Private Sub TestAndBuffer(ByVal oPiece As Element, _
                          ByRef zones() As Element, _
                          ByVal Dist As Double, _
                          ByVal dSlack As Double, _
                          ByRef gStart As Point3d, _
                          ByRef gEnd As Point3d, _
                          ByVal allRound As Boolean, _
                          ByVal tol As Double, _
                          ByRef repBufs() As Element, _
                          ByRef nRep As Long, _
                          ByVal DebugMode As Boolean)
    On Error GoTo ErrorHandler

    Dim dTotal As Double
    Dim dIn    As Double
    Dim buf    As Element

    dTotal = Length.GetLength(oPiece, RndLength:=False)
    If dTotal <= dSlack Then Exit Sub

    dIn = Length.GetPartialLengthInsideZones(oPiece, zones)
    If dIn >= dTotal - dSlack Then Exit Sub

    Select Case oPiece.Type
        Case msdElementTypeLine
            Set buf = BuildLineZone(oPiece, Dist, _
                        CapRoundAt(oPiece.AsChainableElement.StartPoint, gStart, gEnd, allRound, tol), _
                        CapRoundAt(oPiece.AsChainableElement.EndPoint, gStart, gEnd, allRound, tol))
        Case msdElementTypeArc
            Set buf = BuildArcZone(oPiece, Dist, _
                        CapRoundAt(oPiece.AsChainableElement.StartPoint, gStart, gEnd, allRound, tol), _
                        CapRoundAt(oPiece.AsChainableElement.EndPoint, gStart, gEnd, allRound, tol))
    End Select

    If buf Is Nothing Then
        ' The builder itself declined this piece. That is worth saying out loud: it means the hole
        ' was never a merge failure, and no amount of re-merging will close it.
        If DebugMode Then _
            DbgLine "COVER piece type " & oPiece.Type & ", " & Format(dTotal, "0.000") & _
                    " m, " & Format(dIn, "0.000") & " m inside -> NO BUFFER BUILT"
        Exit Sub
    End If

    ReDim Preserve repBufs(0 To nRep)
    Set repBufs(nRep) = buf
    nRep = nRep + 1

    If DebugMode Then _
        DbgLine "COVER piece type " & oPiece.Type & ", " & Format(dTotal, "0.000") & _
                " m, " & Format(dIn, "0.000") & " m inside -> buffer rebuilt"
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.TestAndBuffer"
End Sub

' The summed area of a set of zones, skipping any whose area cannot be read (AreaOf returns a huge
' sentinel for those, which would swamp the total). Used by the DebugMode trace, where weighing the
' repair fusion is what tells a patch the union ABSORBED from one it silently dropped.
Public Function ZonesArea(ByRef els() As Element, ByVal n As Long) As Double
    On Error Resume Next
    Dim k As Long
    Dim d As Double
    For k = 0 To n - 1
        d = AreaOf(els(k))
        If d < 1E+29 Then ZonesArea = ZonesArea + d
    Next k
End Function

' True for the element types the coverage check can judge: Length must be able to measure them end
' to end, AND DispatchElement must build a zone for them. Both halves matter.
'
' A cell or an ellipse reads back as zero length, and zero length against a non-zero total would
' accuse every one of them. A plain Shape measures fine but has no case in DispatchElement, so it
' never receives a zone at all - and that is DELIBERATE (Asketyll, 2026-09-04: "Shape non zonne"),
' not an oversight for this pass to repair. Flagging one would be an accusation nothing could act
' on, repeated on every run.
Private Function Measurable(ByVal oEl As Element) As Boolean
    Select Case oEl.Type
        Case msdElementTypeLine, msdElementTypeLineString, msdElementTypeArc, _
             msdElementTypeComplexString, msdElementTypeComplexShape
            Measurable = True
    End Select
End Function
