' Module: Zoning_Buffer
' Description: The buffer of a chain (LineString / ComplexString) as ONE closed shape of EXACT lines and
'              arcs, built from the definition of a buffer: offset every primitive naively, join the
'              convex corners, then cut away whatever ends up nearer the chain than the offset distance.
'              No GetRegionUnion, no corner algebra. Rationale: _bmad/docs/zoning-mechanics.md
' License: This project is licensed under the AGPL-3.0.
' Dependencies: Geometry, Length, Zoning (AreaOf), ErrorHandlerClass
Option Explicit

' A line or a circular arc, always in WORLD coordinates. An arc's own StartAngle/SweepAngle live in its
' element frame and must never be mixed with world points - everything here is re-derived from the
' centre and the two endpoints, using only the SIGN of the stored sweep.
Public Type ChainSeg
    IsArc    As Boolean
    P0       As Point3d
    P1       As Point3d
    Centre   As Point3d
    Radius   As Double
    StartAng As Double
    Sweep    As Double
End Type

' Sampling step when hunting for the places where the naive contour dives inside the buffer, and the
' angular step of a round join, both as a share of the offset distance. A self-overlap shorter than this
' is not worth cutting.
Private Const STEP_RATIO As Double = 0.02

' A contour point is inside - and must be cut away - when it is nearer the chain than Dist by more than
' this share of Dist. Distances here are exact, so this only has to absorb floating-point noise.
Private Const INSIDE_RATIO As Double = 0.001

' Refuse rather than grind.
Private Const MAX_SEGS As Long = 4000

' The finished shape must contain the element it was built from, to within this share of Dist.
Private Const COVER_SLACK_RATIO As Double = 0.05

' Diagnostics, off in production. Consts rather than variables: a VBA project reset wipes module
' state and the trace goes silent just when it is needed. TRACE logs one line per element through
' Zoning.DbgLine; KEEP_UNSOUND writes a contour that failed IsSound so it can be looked at.
Public Const BUFFER_TRACE As Boolean = False
Public Const BUFFER_KEEP_UNSOUND As Boolean = False

Private msWhy As String
Private mbUnsound As Boolean

' BuildBuffer
' ---------------------------------------------------------------------------
' The buffer of oEl as one closed ComplexShape of exact lines and arcs, or Nothing.
Public Function BuildBuffer(ByVal oEl As Element, _
                            ByVal Dist As Double, _
                            ByVal RoundCaps As Boolean) As Element
    On Error Resume Next

    msWhy = "unknown"
    mbUnsound = False
    Set BuildBuffer = TryBuildBuffer(oEl, Dist, RoundCaps)

    If BUFFER_TRACE Then
        Dim sHead As String
        If BuildBuffer Is Nothing Then
            sHead = "buffer REFUSED (" & msWhy & ")"
        ElseIf mbUnsound Then
            sHead = "buffer KEPT-UNSOUND (" & msWhy & ") - written for inspection"
        Else
            sHead = "buffer OK"
        End If
        ErrorHandler.HandleError sHead & " type=" & CStr(oEl.Type) & _
                                 " dist=" & Format(Dist, "0.000") & " id=" & DLongToString(oEl.ID), _
                                 0, "", "Zoning_Buffer"
    End If
End Function

Private Function TryBuildBuffer(ByVal oEl As Element, _
                                ByVal Dist As Double, _
                                ByVal RoundCaps As Boolean) As Element
    On Error GoTo ErrorHandler

    Dim chain()  As ChainSeg
    Dim nChain   As Long
    Dim rev()    As ChainSeg
    Dim raw()    As ChainSeg
    Dim nRaw     As Long
    Dim keep()   As ChainSeg
    Dim nKeep    As Long
    Dim oShape   As Element

    If Dist <= 0 Then Why "dist <= 0": Exit Function
    If Not Flatten(oEl, chain, nChain) Then Exit Function
    If nChain < 1 Then Why "empty chain": Exit Function

    ' The naive contour: both sides offset (the return side is the left offset of the reversed chain),
    ' convex corners rounded, concave corners left overlapping on purpose, two caps. Nothing can fail
    ' here - there is no corner to solve, only local construction.
    ' Both sides are built FIRST, then the caps: a cap has to run from where one side really ends to
    ' where the other really starts, and those points are not known until both sides exist. Deriving a
    ' cap from an assumed half-turn instead put its far end wherever the arithmetic landed, leaving a
    ' hole that CreateComplexShapeElement1 then "closed" by reversing components - the snail.
    Dim leftS()  As ChainSeg
    Dim nLeft    As Long
    Dim rightS() As ChainSeg
    Dim nRight   As Long

    nLeft = 0
    nRight = 0
    If Not NaiveSide(chain, nChain, Dist, leftS, nLeft) Then Exit Function
    Reverse chain, nChain, rev
    If Not NaiveSide(rev, nChain, Dist, rightS, nRight) Then Exit Function
    If nLeft < 1 Or nRight < 1 Then Why "an offset side is empty": Exit Function

    nRaw = 0
    AppendAll raw, nRaw, leftS, nLeft
    AppendCap raw, nRaw, chain(nChain - 1).P1, leftS(nLeft - 1).P1, rightS(0).P0, Dist, RoundCaps, chain, nChain
    AppendAll raw, nRaw, rightS, nRight
    AppendCap raw, nRaw, chain(0).P0, rightS(nRight - 1).P1, leftS(0).P0, Dist, RoundCaps, chain, nChain

    ' The only correction, and it needs no geometry beyond a distance: whatever lies nearer the chain
    ' than Dist is not on the boundary. Primitives are CUT, not replaced - an arc stays an arc.
    If Not KeepOutside(raw, nRaw, chain, nChain, Dist, keep, nKeep) Then Exit Function
    If nKeep < 3 Then Why "fewer than 3 pieces left": Exit Function

    ' The cut lands on a threshold just inside Dist, and that small radial error is amplified along the
    ' pieces by 1/tan(half-angle) - at a shallow corner it becomes millimetres, and the two ends no
    ' longer meet. Welding each pair back onto the intersection of their own supports puts the sharp
    ' vertex back, which is what a corner of the buffer actually is.
    WeldJoints keep, nKeep, Dist

    Set oShape = Assemble(keep, nKeep)
    If oShape Is Nothing Then Why "assembly refused": Exit Function
    If Not IsSound(oShape, oEl, Dist) Then
        If Not BUFFER_KEEP_UNSOUND Then Exit Function
        mbUnsound = True
    End If

    Set TryBuildBuffer = oShape
    Exit Function

ErrorHandler:
    Why "raised: " & Err.Description
    Set TryBuildBuffer = Nothing
End Function

Public Function LastRefusal() As String
    LastRefusal = msWhy
End Function

Private Sub Why(ByVal s As String)
    msWhy = s
End Sub

'######################################################################################################################
'                                          THE CHAIN
'######################################################################################################################

' Public: also used by GroupSplitGuard to measure how near a group member lies to each piece of a cut.
Public Function Flatten(ByVal oEl As Element, ByRef segs() As ChainSeg, ByRef nSeg As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim oEnum As ElementEnumerator
    Dim comp  As Element

    nSeg = 0
    Select Case oEl.Type
        Case msdElementTypeLine
            AddLine segs, nSeg, oEl.AsChainableElement.StartPoint, oEl.AsChainableElement.EndPoint

        Case msdElementTypeArc
            If Not AddArc(segs, nSeg, oEl) Then Exit Function

        Case msdElementTypeLineString
            If Not AddVertices(oEl, segs, nSeg) Then Exit Function

        Case msdElementTypeComplexString
            Set oEnum = oEl.AsComplexStringElement.GetSubElements
            Do While oEnum.MoveNext
                Set comp = oEnum.Current
                Select Case comp.Type
                    Case msdElementTypeLine
                        AddLine segs, nSeg, comp.AsChainableElement.StartPoint, comp.AsChainableElement.EndPoint
                    Case msdElementTypeLineString
                        If Not AddVertices(comp, segs, nSeg) Then Exit Function
                    Case msdElementTypeArc
                        If Not AddArc(segs, nSeg, comp) Then Exit Function
                    Case Else
                        Why "unknown sub-element type " & CStr(comp.Type): Exit Function
                End Select
                If nSeg > MAX_SEGS Then Why "chain too long": Exit Function
            Loop

        Case Else
            Why "not a chain type": Exit Function
    End Select

    Flatten = (nSeg > 0)
    Exit Function

ErrorHandler:
    Flatten = False
End Function

Private Function AddVertices(ByVal oEl As Element, ByRef segs() As ChainSeg, ByRef nSeg As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim oVL As VertexList
    Dim v() As Point3d
    Dim j   As Long

    Set oVL = oEl
    v = oVL.GetVertices
    For j = LBound(v) To UBound(v) - 1
        AddLine segs, nSeg, v(j), v(j + 1)
    Next j

    AddVertices = True
    Exit Function

ErrorHandler:
    AddVertices = False
End Function

' Zero-length segments are dropped: their direction is undefined and every corner test would inherit it.
Private Sub AddLine(ByRef segs() As ChainSeg, ByRef nSeg As Long, ByRef p0 As Point3d, ByRef p1 As Point3d)
    If Point3dDistanceXY(p0, p1) < 0.000001 Then Exit Sub
    ReDim Preserve segs(0 To nSeg)
    segs(nSeg).IsArc = False
    segs(nSeg).P0 = p0
    segs(nSeg).P1 = p1
    nSeg = nSeg + 1
End Sub

Private Function AddArc(ByRef segs() As ChainSeg, ByRef nSeg As Long, ByVal oEl As Element) As Boolean
    On Error GoTo ErrorHandler

    Dim a    As ArcElement
    Dim ptC  As Point3d
    Dim r    As Double
    Dim angS As Double

    Set a = oEl
    ptC = a.CenterPoint
    r = Point3dDistanceXY(a.StartPoint, ptC)
    If r < 0.000001 Then Why "degenerate arc": Exit Function
    If Abs(Point3dDistanceXY(a.EndPoint, ptC) - r) > r * 0.0001 Then Why "arc not circular": Exit Function

    angS = Point3dPolarAngle(Point3dSubtract(a.StartPoint, ptC))

    ReDim Preserve segs(0 To nSeg)
    segs(nSeg).IsArc = True
    segs(nSeg).Centre = ptC
    segs(nSeg).Radius = r
    segs(nSeg).StartAng = angS
    segs(nSeg).Sweep = Geometry.NormalizeAngle( _
        Point3dPolarAngle(Point3dSubtract(a.EndPoint, ptC)) - angS, a.SweepAngle)
    segs(nSeg).P0 = a.StartPoint
    segs(nSeg).P1 = a.EndPoint
    nSeg = nSeg + 1

    AddArc = True
    Exit Function

ErrorHandler:
    AddArc = False
End Function

Private Sub Reverse(ByRef segs() As ChainSeg, ByVal nSeg As Long, ByRef out() As ChainSeg)
    Dim i As Long
    ReDim out(0 To nSeg - 1)
    For i = 0 To nSeg - 1
        out(i) = segs(nSeg - 1 - i)
        out(i).P0 = segs(nSeg - 1 - i).P1
        out(i).P1 = segs(nSeg - 1 - i).P0
        If out(i).IsArc Then
            out(i).StartAng = segs(nSeg - 1 - i).StartAng + segs(nSeg - 1 - i).Sweep
            out(i).Sweep = -segs(nSeg - 1 - i).Sweep
        End If
    Next i
End Sub

'######################################################################################################################
'                                          THE NAIVE CONTOUR
'######################################################################################################################

' One side, offset LEFT, appended to the contour being built. Convex corners get a round join; concave
' corners get nothing at all and are left to overlap - the cut pass deals with them.
Private Function NaiveSide(ByRef chain() As ChainSeg, ByVal nChain As Long, ByVal Dist As Double, _
                           ByRef out() As ChainSeg, ByRef nOut As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim i     As Long
    Dim o     As ChainSeg
    Dim turn  As Double
    Dim ptEnd As Point3d

    For i = 0 To nChain - 1
        If OffsetSeg(chain(i), Dist, o) Then
            If nOut > 0 And i > 0 Then
                turn = SignedTurn(Tangent(chain(i - 1), True), Tangent(chain(i), False))
                If turn < 0 Then
                    ptEnd = out(nOut - 1).P1
                    AppendArcFan out, nOut, chain(i).P0, ptEnd, o.P0, Dist
                End If
            End If
            Append out, nOut, o
            If nOut > MAX_SEGS Then Why "contour too long": Exit Function
        End If
    Next i

    NaiveSide = True
    Exit Function

ErrorHandler:
    NaiveSide = False
End Function

' A primitive offset to its left: a line by its normal, an arc by its radius alone.
Private Function OffsetSeg(ByRef seg As ChainSeg, ByVal Dist As Double, ByRef out As ChainSeg) As Boolean
    On Error GoTo ErrorHandler

    Dim n As Point3d
    Dim r As Double

    out = seg
    If Not seg.IsArc Then
        n = Geometry.Perp2D(seg.P0, seg.P1, Dist)
        If Abs(n.X) + Abs(n.Y) < 0.000000001 Then Exit Function
        out.P0 = Point3dAdd(seg.P0, n)
        out.P1 = Point3dAdd(seg.P1, n)
        OffsetSeg = True
        Exit Function
    End If

    ' Counter-clockwise the centre is on the left, so the left offset shrinks the radius.
    r = seg.Radius - Sgn(seg.Sweep) * Dist
    If r <= Dist * 0.001 Then Exit Function      ' the bend is tighter than the offset: no left arc
    out.Radius = r
    out.P0 = PointAt(seg.Centre, r, seg.StartAng)
    out.P1 = PointAt(seg.Centre, r, seg.StartAng + seg.Sweep)
    OffsetSeg = True
    Exit Function

ErrorHandler:
    OffsetSeg = False
End Function

' The round join filling the outside of a turn: ONE arc of radius Dist about the vertex. Exact, so the
' zone keeps a true curve where the cable turns.
Private Sub AppendArcFan(ByRef out() As ChainSeg, ByRef nOut As Long, ByRef ptC As Point3d, _
                         ByRef ptFrom As Point3d, ByRef ptTo As Point3d, ByVal Dist As Double)
    On Error GoTo ErrorHandler

    Dim s  As ChainSeg
    Dim a0 As Double

    a0 = Point3dPolarAngle(Point3dSubtract(ptFrom, ptC))
    s.IsArc = True
    s.Centre = ptC
    s.Radius = Dist
    s.StartAng = a0
    s.Sweep = SignedTurn(Point3dSubtract(ptFrom, ptC), Point3dSubtract(ptTo, ptC))
    If Abs(s.Sweep) < 0.000001 Then Exit Sub
    s.P0 = ptFrom
    s.P1 = ptTo
    Append out, nOut, s
    Exit Sub

ErrorHandler:
End Sub

' The cap at a free end: an arc of radius Dist about the end point, running from where one side stops
' to where the other starts. A flat cap adds nothing - Assemble closes it with a straight edge.
' The sweep DIRECTION is not assumed: both are tried and the one whose midpoint sits farther from the
' chain wins. That is what a cap is - the way round that goes AWAY from the cable - and deciding it by
' measurement rather than by a sign convention is what stops the arc wrapping the wrong way.
Private Sub AppendCap(ByRef out() As ChainSeg, ByRef nOut As Long, ByRef ptEnd As Point3d, _
                      ByRef ptFrom As Point3d, ByRef ptTo As Point3d, ByVal Dist As Double, _
                      ByVal bRound As Boolean, ByRef chain() As ChainSeg, ByVal nChain As Long)
    On Error GoTo ErrorHandler

    Dim s     As ChainSeg
    Dim a0    As Double
    Dim swCCW As Double
    Dim swCW  As Double
    Dim dCCW  As Double
    Dim dCW   As Double

    If Not bRound Then Exit Sub

    a0 = Point3dPolarAngle(Point3dSubtract(ptFrom, ptEnd))
    swCCW = Geometry.NormalizeAngle(Point3dPolarAngle(Point3dSubtract(ptTo, ptEnd)) - a0, 1)
    swCW = Geometry.NormalizeAngle(Point3dPolarAngle(Point3dSubtract(ptTo, ptEnd)) - a0, -1)

    dCCW = DistanceToChain(PointAt(ptEnd, Dist, a0 + swCCW / 2), chain, nChain)
    dCW = DistanceToChain(PointAt(ptEnd, Dist, a0 + swCW / 2), chain, nChain)

    s.IsArc = True
    s.Centre = ptEnd
    s.Radius = Dist
    s.StartAng = a0
    If dCCW >= dCW Then s.Sweep = swCCW Else s.Sweep = swCW
    If Abs(s.Sweep) < 0.000001 Then Exit Sub
    s.P0 = ptFrom
    s.P1 = ptTo
    Append out, nOut, s
    Exit Sub

ErrorHandler:
End Sub

Private Sub AppendAll(ByRef out() As ChainSeg, ByRef nOut As Long, ByRef src() As ChainSeg, ByVal n As Long)
    Dim i As Long
    For i = 0 To n - 1
        Append out, nOut, src(i)
    Next i
End Sub

'######################################################################################################################
'                                          THE CUT
'######################################################################################################################

' Walks every contour primitive and keeps only the stretches that really are Dist away from the chain.
' A primitive is CUT, never replaced: a trimmed line is a line, a trimmed arc is an arc with a smaller
' sweep. The transitions are found by sampling and then refined by bisection, so the cut lands on the
' self-crossing itself rather than on the nearest sample.
Private Function KeepOutside(ByRef raw() As ChainSeg, ByVal nRaw As Long, _
                             ByRef chain() As ChainSeg, ByVal nChain As Long, ByVal Dist As Double, _
                             ByRef keep() As ChainSeg, ByRef nKeep As Long) As Boolean
    On Error GoTo ErrorHandler

    Dim i      As Long
    Dim k      As Long
    Dim nStep  As Long
    Dim t      As Double
    Dim tPrev  As Double
    Dim bIn    As Boolean
    Dim bPrev  As Boolean
    Dim tStart As Double
    Dim lim    As Double
    Dim piece  As ChainSeg

    lim = Dist * (1 - INSIDE_RATIO)
    nKeep = 0

    For i = 0 To nRaw - 1
        nStep = SampleCount(raw(i), Dist)
        bPrev = (DistanceToChain(AtParam(raw(i), 0), chain, nChain) >= lim)
        tPrev = 0
        tStart = 0

        For k = 1 To nStep
            t = k / nStep
            bIn = (DistanceToChain(AtParam(raw(i), t), chain, nChain) >= lim)

            If bIn <> bPrev Then
                ' Refine where the boundary is actually crossed, then open or close a kept stretch.
                t = Bisect(raw(i), chain, nChain, tPrev, t, lim)
                If bIn Then
                    tStart = t
                Else
                    If SubPiece(raw(i), tStart, t, piece) Then Append keep, nKeep, piece
                End If
                bPrev = bIn
            End If
            tPrev = t
        Next k

        If bPrev Then
            If SubPiece(raw(i), tStart, 1, piece) Then Append keep, nKeep, piece
        End If
        If nKeep > MAX_SEGS Then Why "too many pieces": Exit Function
    Next i

    KeepOutside = True
    Exit Function

ErrorHandler:
    KeepOutside = False
End Function

' Where between t0 and t1 the distance crosses lim. Twenty halvings put it well under a micron.
Private Function Bisect(ByRef seg As ChainSeg, ByRef chain() As ChainSeg, ByVal nChain As Long, _
                        ByVal t0 As Double, ByVal t1 As Double, ByVal lim As Double) As Double
    Dim lo As Double
    Dim hi As Double
    Dim mid As Double
    Dim k  As Long
    Dim bLo As Boolean

    lo = t0
    hi = t1
    bLo = (DistanceToChain(AtParam(seg, lo), chain, nChain) >= lim)

    For k = 1 To 20
        mid = (lo + hi) / 2
        If (DistanceToChain(AtParam(seg, mid), chain, nChain) >= lim) = bLo Then
            lo = mid
        Else
            hi = mid
        End If
    Next k

    Bisect = (lo + hi) / 2
End Function

' The exact distance from a point to the chain. No densification: a line uses the bounded-ray closest
' point, an arc its own radius when the point projects inside the sweep and its endpoints otherwise.
Public Function DistanceToChain(ByRef p As Point3d, ByRef chain() As ChainSeg, ByVal nChain As Long) As Double
    On Error GoTo ErrorHandler

    Dim i    As Long
    Dim best As Double
    Dim d    As Double

    best = 1E+30
    For i = 0 To nChain - 1
        If Not FarFrom(p, chain(i), best) Then
            d = DistanceToSeg(p, chain(i))
            If d < best Then best = d
        End If
    Next i

    DistanceToChain = best
    Exit Function

ErrorHandler:
    DistanceToChain = 1E+30
End Function

Private Function DistanceToSeg(ByRef p As Point3d, ByRef seg As ChainSeg) As Double
    On Error GoTo ErrorHandler

    Dim ray     As Ray3d
    Dim ptClose As Point3d
    Dim frac    As Double
    Dim d       As Double
    Dim ang     As Double
    Dim t       As Double
    Dim dEnd    As Double

    If Not seg.IsArc Then
        ray = Ray3dFromPoint3dStartEnd(seg.P0, seg.P1)
        Ray3dClosestPointBoundedXY ray, p, ptClose, frac
        DistanceToSeg = Point3dDistanceXY(p, ptClose)
        Exit Function
    End If

    d = Point3dDistanceXY(p, seg.Centre)
    If d > 0.000001 Then
        ang = Point3dPolarAngle(Point3dSubtract(p, seg.Centre))
        t = Geometry.NormalizeAngle(ang - seg.StartAng, Sgn(seg.Sweep))
        If Abs(t) <= Abs(seg.Sweep) Then
            DistanceToSeg = Abs(d - seg.Radius)
            Exit Function
        End If
    End If

    DistanceToSeg = Point3dDistanceXY(p, seg.P0)
    dEnd = Point3dDistanceXY(p, seg.P1)
    If dEnd < DistanceToSeg Then DistanceToSeg = dEnd
    Exit Function

ErrorHandler:
    DistanceToSeg = 1E+30
End Function

' Bounding-box rejection: four comparisons dismiss nearly every primitive before any real work.
Private Function FarFrom(ByRef p As Point3d, ByRef seg As ChainSeg, ByVal best As Double) As Boolean
    Dim lo As Double
    Dim hi As Double
    Dim r  As Double

    If seg.IsArc Then r = seg.Radius

    If seg.IsArc Then
        lo = seg.Centre.X - r: hi = seg.Centre.X + r
    ElseIf seg.P0.X < seg.P1.X Then
        lo = seg.P0.X: hi = seg.P1.X
    Else
        lo = seg.P1.X: hi = seg.P0.X
    End If
    If p.X < lo - best Then FarFrom = True: Exit Function
    If p.X > hi + best Then FarFrom = True: Exit Function

    If seg.IsArc Then
        lo = seg.Centre.Y - r: hi = seg.Centre.Y + r
    ElseIf seg.P0.Y < seg.P1.Y Then
        lo = seg.P0.Y: hi = seg.P1.Y
    Else
        lo = seg.P1.Y: hi = seg.P0.Y
    End If
    If p.Y < lo - best Then FarFrom = True: Exit Function
    If p.Y > hi + best Then FarFrom = True: Exit Function
End Function

'######################################################################################################################
'                                          WELDING THE JOINTS
'######################################################################################################################

' Pulls each pair of consecutive pieces back onto the intersection of their supports. Only pairs that
' already nearly meet are touched, and only when the intersection SHORTENS both - so a genuine gap (two
' pieces that belong apart) is left alone for Assemble to bridge.
Private Sub WeldJoints(ByRef segs() As ChainSeg, ByVal nSeg As Long, ByVal Dist As Double)
    On Error GoTo ErrorHandler

    Dim i    As Long
    Dim gap  As Double
    Dim ptX  As Point3d
    Dim ptRef As Point3d
    Dim a    As ChainSeg
    Dim b    As ChainSeg

    For i = 1 To nSeg - 1
        ' No upper bound on the gap: at a shallow corner the recovered vertex is legitimately far, and
        ' a whole piece swallowed by a bend leaves a wide hole whose vertex is farther still. What keeps
        ' this safe is not a distance limit but the rule that the weld must SHORTEN both pieces.
        gap = Point3dDistanceXY(segs(i - 1).P1, segs(i).P0)
        If gap > 0.000001 Then
            ptRef = Point3dFromXYZ((segs(i - 1).P1.X + segs(i).P0.X) / 2, _
                                   (segs(i - 1).P1.Y + segs(i).P0.Y) / 2, segs(i - 1).P1.Z)
            If Intersect(segs(i - 1), segs(i), ptRef, ptX) Then
                a = segs(i - 1)
                b = segs(i)
                If TrimEnd(a, ptX) Then
                    If TrimStart(b, ptX) Then
                        segs(i - 1) = a
                        segs(i) = b
                    End If
                End If
            End If
        End If
    Next i
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_Buffer.WeldJoints"
End Sub

' Moves a piece's end onto pt. False when that would lengthen it - the wrong intersection root.
Private Function TrimEnd(ByRef seg As ChainSeg, ByRef pt As Point3d) As Boolean
    Dim newSweep As Double
    Dim dir      As Point3d

    If Not seg.IsArc Then
        dir = Point3dSubtract(seg.P1, seg.P0)
        If (pt.X - seg.P0.X) * dir.X + (pt.Y - seg.P0.Y) * dir.Y <= 0 Then Exit Function
        seg.P1 = pt
        TrimEnd = True
        Exit Function
    End If

    newSweep = Geometry.NormalizeAngle( _
        Point3dPolarAngle(Point3dSubtract(pt, seg.Centre)) - seg.StartAng, Sgn(seg.Sweep))
    If Abs(newSweep) > Abs(seg.Sweep) + 0.000001 Then Exit Function
    seg.P1 = pt
    seg.Sweep = newSweep
    TrimEnd = True
End Function

Private Function TrimStart(ByRef seg As ChainSeg, ByRef pt As Point3d) As Boolean
    Dim aEnd     As Double
    Dim newStart As Double
    Dim newSweep As Double
    Dim dir      As Point3d

    If Not seg.IsArc Then
        dir = Point3dSubtract(seg.P1, seg.P0)
        If (seg.P1.X - pt.X) * dir.X + (seg.P1.Y - pt.Y) * dir.Y <= 0 Then Exit Function
        seg.P0 = pt
        TrimStart = True
        Exit Function
    End If

    aEnd = seg.StartAng + seg.Sweep
    newStart = Point3dPolarAngle(Point3dSubtract(pt, seg.Centre))
    newSweep = Geometry.NormalizeAngle(aEnd - newStart, Sgn(seg.Sweep))
    If Abs(newSweep) > Abs(seg.Sweep) + 0.000001 Then Exit Function
    seg.StartAng = newStart
    seg.Sweep = newSweep
    seg.P0 = pt
    TrimStart = True
End Function

' Where two pieces' INFINITE supports cross - an endless line for a segment, a whole circle for an arc.
' Infinite on purpose: the vertex being recovered lies just outside where the pieces currently stop.
' The root nearest ptRef wins.
Private Function Intersect(ByRef a As ChainSeg, ByRef b As ChainSeg, _
                           ByRef ptRef As Point3d, ByRef ptOut As Point3d) As Boolean
    On Error GoTo ErrorHandler

    Dim cand(0 To 3) As Point3d
    Dim n     As Long
    Dim i     As Long
    Dim d     As Double
    Dim dBest As Double

    If Not a.IsArc And Not b.IsArc Then
        n = XLineLine(a, b, cand)
    ElseIf a.IsArc And b.IsArc Then
        n = XCircleCircle(a, b, cand)
    ElseIf a.IsArc Then
        n = XLineCircle(b, a, cand)
    Else
        n = XLineCircle(a, b, cand)
    End If
    If n = 0 Then Exit Function

    dBest = 1E+30
    For i = 0 To n - 1
        d = Point3dDistanceXY(cand(i), ptRef)
        If d < dBest Then dBest = d: ptOut = cand(i)
    Next i

    Intersect = True
    Exit Function

ErrorHandler:
    Intersect = False
End Function

Private Function XLineLine(ByRef a As ChainSeg, ByRef b As ChainSeg, ByRef cand() As Point3d) As Long
    On Error GoTo ErrorHandler

    Dim u   As Point3d
    Dim v   As Point3d
    Dim den As Double
    Dim t   As Double

    u = Point3dNormalize(Point3dSubtract(a.P1, a.P0))
    v = Point3dNormalize(Point3dSubtract(b.P1, b.P0))
    den = u.X * v.Y - u.Y * v.X
    If Abs(den) < 0.0001 Then Exit Function          ' parallel: nothing to weld onto

    t = ((b.P0.X - a.P0.X) * v.Y - (b.P0.Y - a.P0.Y) * v.X) / den
    cand(0) = Point3dFromXYZ(a.P0.X + t * u.X, a.P0.Y + t * u.Y, a.P0.Z)
    XLineLine = 1
    Exit Function

ErrorHandler:
    XLineLine = 0
End Function

Private Function XLineCircle(ByRef ln As ChainSeg, ByRef arc As ChainSeg, ByRef cand() As Point3d) As Long
    On Error GoTo ErrorHandler

    Dim u    As Point3d
    Dim dx   As Double
    Dim dy   As Double
    Dim bb   As Double
    Dim cc   As Double
    Dim disc As Double
    Dim sq   As Double

    u = Point3dNormalize(Point3dSubtract(ln.P1, ln.P0))
    dx = ln.P0.X - arc.Centre.X
    dy = ln.P0.Y - arc.Centre.Y
    bb = dx * u.X + dy * u.Y
    cc = dx * dx + dy * dy - arc.Radius * arc.Radius
    disc = bb * bb - cc
    If disc < 0 Then Exit Function

    sq = Sqr(disc)
    cand(0) = Point3dFromXYZ(ln.P0.X + (-bb + sq) * u.X, ln.P0.Y + (-bb + sq) * u.Y, ln.P0.Z)
    XLineCircle = 1
    If sq > 0.000001 Then
        cand(1) = Point3dFromXYZ(ln.P0.X + (-bb - sq) * u.X, ln.P0.Y + (-bb - sq) * u.Y, ln.P0.Z)
        XLineCircle = 2
    End If
    Exit Function

ErrorHandler:
    XLineCircle = 0
End Function

Private Function XCircleCircle(ByRef a As ChainSeg, ByRef b As ChainSeg, ByRef cand() As Point3d) As Long
    On Error GoTo ErrorHandler

    Dim dx   As Double
    Dim dy   As Double
    Dim dLen As Double
    Dim aa   As Double
    Dim h2   As Double
    Dim h    As Double
    Dim mx   As Double
    Dim my   As Double

    dx = b.Centre.X - a.Centre.X
    dy = b.Centre.Y - a.Centre.Y
    dLen = Sqr(dx * dx + dy * dy)
    If dLen < 0.000001 Then Exit Function
    If dLen > a.Radius + b.Radius Then Exit Function
    If dLen < Abs(a.Radius - b.Radius) Then Exit Function

    aa = (a.Radius * a.Radius - b.Radius * b.Radius + dLen * dLen) / (2 * dLen)
    h2 = a.Radius * a.Radius - aa * aa
    If h2 < 0 Then h2 = 0
    h = Sqr(h2)
    mx = a.Centre.X + aa * dx / dLen
    my = a.Centre.Y + aa * dy / dLen

    cand(0) = Point3dFromXYZ(mx + h * dy / dLen, my - h * dx / dLen, a.Centre.Z)
    XCircleCircle = 1
    If h > 0.000001 Then
        cand(1) = Point3dFromXYZ(mx - h * dy / dLen, my + h * dx / dLen, a.Centre.Z)
        XCircleCircle = 2
    End If
    Exit Function

ErrorHandler:
    XCircleCircle = 0
End Function

'######################################################################################################################
'                                          PARAMETERS AND PIECES
'######################################################################################################################

Private Function AtParam(ByRef seg As ChainSeg, ByVal t As Double) As Point3d
    If seg.IsArc Then
        AtParam = PointAt(seg.Centre, seg.Radius, seg.StartAng + seg.Sweep * t)
    Else
        AtParam = Point3dFromXYZ(seg.P0.X + (seg.P1.X - seg.P0.X) * t, _
                                 seg.P0.Y + (seg.P1.Y - seg.P0.Y) * t, seg.P0.Z)
    End If
End Function

' The stretch of seg between t0 and t1, still a line or still an arc. False when it is too short to draw.
Private Function SubPiece(ByRef seg As ChainSeg, ByVal t0 As Double, ByVal t1 As Double, _
                          ByRef out As ChainSeg) As Boolean
    If t1 - t0 < 0.000001 Then Exit Function

    out = seg
    out.P0 = AtParam(seg, t0)
    out.P1 = AtParam(seg, t1)
    If seg.IsArc Then
        out.StartAng = seg.StartAng + seg.Sweep * t0
        out.Sweep = seg.Sweep * (t1 - t0)
        If Abs(out.Sweep) < 0.000001 Then Exit Function
    Else
        If Point3dDistanceXY(out.P0, out.P1) < 0.000001 Then Exit Function
    End If
    SubPiece = True
End Function

Private Function SampleCount(ByRef seg As ChainSeg, ByVal Dist As Double) As Long
    Dim L As Double
    If seg.IsArc Then L = Abs(seg.Sweep) * seg.Radius Else L = Point3dDistanceXY(seg.P0, seg.P1)
    SampleCount = Int(L / (Dist * STEP_RATIO)) + 1
    If SampleCount < 2 Then SampleCount = 2
    If SampleCount > 2000 Then SampleCount = 2000
End Function

'######################################################################################################################
'                                          ASSEMBLY AND VALIDATION
'######################################################################################################################

' Kept pieces in order, with a straight edge wherever a cut left a gap. CreateComplexShapeElement1
' reverses components as needed to close the loop.
Private Function Assemble(ByRef segs() As ChainSeg, ByVal nSeg As Long) As Element
    On Error GoTo ErrorHandler

    Dim comps() As ChainableElement
    Dim nComp   As Long
    Dim i       As Long
    Dim oEl     As Element

    ReDim comps(0 To nSeg * 2)
    nComp = 0

    For i = 0 To nSeg - 1
        If i > 0 Then
            If Point3dDistanceXY(segs(i - 1).P1, segs(i).P0) > 0.000001 Then
                Set oEl = CreateLineElement2(Nothing, segs(i - 1).P1, segs(i).P0)
                Set comps(nComp) = oEl: nComp = nComp + 1
            End If
        End If
        Set oEl = ToElement(segs(i))
        If oEl Is Nothing Then Exit Function
        Set comps(nComp) = oEl: nComp = nComp + 1
    Next i

    ' Close the loop.
    If Point3dDistanceXY(segs(nSeg - 1).P1, segs(0).P0) > 0.000001 Then
        Set oEl = CreateLineElement2(Nothing, segs(nSeg - 1).P1, segs(0).P0)
        Set comps(nComp) = oEl: nComp = nComp + 1
    End If

    If nComp < 3 Then Exit Function
    ReDim Preserve comps(0 To nComp - 1)
    Set Assemble = CreateComplexShapeElement1(comps, msdFillModeNotFilled)
    Exit Function

ErrorHandler:
    Set Assemble = Nothing
End Function

Private Function ToElement(ByRef seg As ChainSeg) As Element
    On Error GoTo ErrorHandler
    If seg.IsArc Then
        Set ToElement = CreateArcElement2(Nothing, seg.Centre, seg.Radius, seg.Radius, _
                                          Matrix3dIdentity, seg.StartAng, seg.Sweep)
    Else
        Set ToElement = CreateLineElement2(Nothing, seg.P0, seg.P1)
    End If
    Exit Function
ErrorHandler:
    Set ToElement = Nothing
End Function

Private Function IsSound(ByVal oShape As Element, ByVal oSrc As Element, ByVal Dist As Double) As Boolean
    On Error GoTo ErrorHandler

    Dim dArea   As Double
    Dim zones(0 To 0) As Element
    Dim dInside As Double
    Dim dTotal  As Double

    dArea = Zoning.AreaOf(oShape)
    If dArea <= 0 Or dArea > 1E+29 Then Why "no usable area": Exit Function

    dTotal = Length.GetLength(oSrc, RndLength:=False)
    If dTotal <= 0 Then IsSound = True: Exit Function

    Set zones(0) = oShape
    dInside = Length.GetPartialLengthInsideZones(oSrc, zones)
    IsSound = (dInside >= dTotal - Dist * COVER_SLACK_RATIO)
    If Not IsSound Then Why "coverage " & Format(dInside, "0.000") & " of " & Format(dTotal, "0.000")
    Exit Function

ErrorHandler:
    IsSound = False
End Function

'######################################################################################################################
'                                          SMALL HELPERS
'######################################################################################################################

' The unit tangent of a primitive in the direction of travel, at its start or its end.
Private Function Tangent(ByRef seg As ChainSeg, ByVal bAtEnd As Boolean) As Point3d
    On Error GoTo ErrorHandler

    Dim ang As Double

    If Not seg.IsArc Then
        Tangent = Point3dNormalize(Point3dSubtract(seg.P1, seg.P0))
        Exit Function
    End If

    ang = seg.StartAng
    If bAtEnd Then ang = seg.StartAng + seg.Sweep
    Tangent = Point3dNormalize(Point3dFromXYZ(-Sgn(seg.Sweep) * Sin(ang), Sgn(seg.Sweep) * Cos(ang), 0))
    Exit Function

ErrorHandler:
End Function

' Signed angle from v1 to v2 about +Z, in [-pi, pi]. Negative = turning right.
' Vector3d is a DIFFERENT type from Point3d in the type library, hence the explicit conversions.
Private Function SignedTurn(ByRef v1 As Point3d, ByRef v2 As Point3d) As Double
    On Error GoTo ErrorHandler

    Dim a As Vector3d
    Dim b As Vector3d
    Dim z As Vector3d

    a = Vector3dFromXYZ(v1.X, v1.Y, 0)
    b = Vector3dFromXYZ(v2.X, v2.Y, 0)
    z = Vector3dFromXYZ(0, 0, 1)
    SignedTurn = Vector3dSignedAngleBetweenVectors(a, b, z)
    Exit Function

ErrorHandler:
    SignedTurn = 0
End Function

Private Function PointAt(ByRef ptC As Point3d, ByVal r As Double, ByVal ang As Double) As Point3d
    PointAt = Point3dAdd(ptC, Point3dFromXYZ(r * Cos(ang), r * Sin(ang), 0))
End Function

Private Sub Append(ByRef out() As ChainSeg, ByRef nOut As Long, ByRef seg As ChainSeg)
    ReDim Preserve out(0 To nOut)
    out(nOut) = seg
    nOut = nOut + 1
End Sub
