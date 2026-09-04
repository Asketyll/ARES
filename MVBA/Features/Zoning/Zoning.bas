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
' Rationale, thresholds and the measurements behind them: _bmad/docs/zoning-mechanics.md
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, ARESConstants, ErrorHandlerClass, Geometry, GetElements

Option Explicit

' Fusion trace (DebugMode only): the file keeps every line, the Immediate window shows the first ones.
Private Const DBG_FILE As String = "C:\ARES\ARES_zoning_debug.log"
Private Const DBG_ECHO_MAX As Long = 120
Private mnDbgShown As Long

' Generates offset zones around elements on the specified source levels.
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

    ' Read once here rather than per zone: Val, not CDbl, because ARES stores dot-decimals and CDbl
    ' is locale-aware - on a French install CDbl("0.5") is 5, which would multiply every cleanup
    ' threshold by ten. A missing or unreadable value means 1, the normal behaviour.
    Dim dCleanup As Double
    dCleanup = Val(Replace(ARESConfig.ARES_ZONING_CLEANUP_FACTOR.Value, ",", "."))
    If Len(Trim(ARESConfig.ARES_ZONING_CLEANUP_FACTOR.Value)) = 0 Then dCleanup = 1
    Zoning_Cleanup.SetCleanupFactor dCleanup
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
        DispatchElement Elements(i), Dist, TargetLevel, Color, Style, Weight, _
                        allBufs, nAllBufs, DebugMode, RoundCaps
    Next i

    ' --- Merge all accumulated zones and write to the model (MergeZones = True only) ---
    If MergeZones And nAllBufs > 0 Then
        ' Debug mode: write a clone of each pre-merge shape so the individual zones
        ' are visible alongside the final merged result.
        If DebugMode Then WriteDebugClones allBufs, nAllBufs, TargetLevel, Color, Style, Weight

        Dim mergedAll() As Element
        Dim nMergedAll  As Long
        FuseRegions allBufs, nAllBufs, mergedAll, nMergedAll, DebugMode, "global"

        ' --- Last word: nothing of a cable may end up outside the zones ---
        ' The union is where coverage gets lost - a buffer that fails to merge is simply absent from
        ' the result, and nothing downstream would ever notice. So the source elements are measured
        ' against what the union actually produced, and whatever is missing is rebuilt and merged in.
        Dim repBufs() As Element
        Dim nRep      As Long
        nRep = 0
        RepairUncovered Elements, mergedAll, nMergedAll, Dist, TargetLevel, Color, Style, Weight, _
                        DebugMode, RoundCaps, repBufs, nRep

        If nRep > 0 Then
            ' Fused WITH the zones rather than beside them: a gap in the middle of a cable can bridge
            ' two zones that were never joined, or fill a hole inside one, and only the union knows
            ' which. Feeding the zones back in is what lets either happen.
            Dim again()  As Element
            Dim nAgain   As Long
            ReDim again(0 To nMergedAll + nRep - 1)
            For k = 0 To nMergedAll - 1
                Set again(k) = mergedAll(k)
            Next k
            For k = 0 To nRep - 1
                Set again(nMergedAll + k) = repBufs(k)
            Next k
            nAgain = nMergedAll + nRep

            Dim nBefore As Long
            Dim dBefore As Double
            Dim dPatch  As Double
            nBefore = nMergedAll
            dBefore = ZonesArea(mergedAll, nMergedAll)
            dPatch = ZonesArea(repBufs, nRep)
            FuseRegions again, nAgain, mergedAll, nMergedAll, DebugMode, "coverage repair"

            ' The decisive line, and the one that was missing: a patch can be correct, be handed to
            ' the union, and still come back as its own separate shape. Fewer zones out than in means
            ' it welded; the same count or more means the union refused it, and no better patch will
            ' change that.
            ' Area is what separates "absorbed but did not bridge" from "silently dropped", and the
            ' two call for opposite fixes. A union that swallowed the patches grows by most of their
            ' area; one that ignored them comes back the same size it went in.
            If DebugMode Then _
                DbgLine "COVER fusion: " & nBefore & " zone(s) " & Format(dBefore, "0.00") & " m2" & _
                        " + " & nRep & " patch(es) " & Format(dPatch, "0.00") & " m2" & _
                        " -> " & nMergedAll & " zone(s) " & Format(ZonesArea(mergedAll, nMergedAll), "0.00") & " m2"
        End If

        For k = 0 To nMergedAll - 1
            WriteEl mergedAll(k), TargetLevel, Color, Style, Weight, Dist
        Next k
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning.Zoning"
End Sub

' ============================================================
'  OUTPUT HELPERS
' ============================================================

' DbgLine
' ---------------------------------------------------------------------------
' Debug output for the fusion trace. Goes to a FILE first and to the Immediate window only for the
' first lines of each run: a drawing that folds a thousand buffers floods the Immediate window's
' buffer and pushes the beginning - where the useful part is - out of reach.
Public Sub DbgLine(ByVal sMsg As String)
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
' The input buffers are MOVED IN PLACE (near origin, a precision workaround); callers must not
' reuse bufs() afterwards. nBuf = 1 is returned as-is, no union.
Public Sub FuseRegions(ByRef bufs() As Element, _
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

' WriteDebugClones
' ---------------------------------------------------------------------------
' Writes a clone of each pre-merge buffer to the model so the individual zones are visible
' alongside the final merged result (DebugMode only).
Public Sub WriteDebugClones(ByRef bufs() As Element, _
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
Public Sub AddOrWrite(ByVal oEl As Element, _
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