' ExportLengthInRegion.bas
' Description: Scans graphical elements in the active model, computes the length of
' each element (or partial length) that lies inside any zone produced by Zoning.bas,
' and writes a per-group summary to a new Excel workbook.
'
' SUPPORTED ZONE TYPES   ShapeElement, ComplexShapeElement, EllipseElement,
'                        CellHeader (donut / grouped-holes — outer+inner boundary)
' SUPPORTED CANDIDATES   Line, Arc, LineString, Shape, ComplexString, ComplexShape
'
' PARTIAL-LENGTH STRATEGY  (GetPartialLengthInsideZones — returns 0 when no overlap)
'   All types: one GetIntersectionPoints call per zone.
'     nPts = 0 → start-point check (fully inside → full length; outside → 0).
'               Closed candidates also check reverse containment via zone bbox-center.
'     nPts > 0 → alternate inside/outside walk from sorted crossing points.
'   ComplexString / ComplexShape: fast path at complex level; sub-element decomposition
'     only when nPts > 0.  Shape: segment walk + closing segment.
'
' KNOWN LIMITATION  Reverse containment uses the zone's bbox-center as representative
'   point. Heavily non-convex zones whose geometric centroid falls outside the boundary
'   (e.g. a C-shaped zone) may fail this check. Standard convex/rectangular/elliptical
'   zones are handled correctly.
'
' EXCEL COM CONTRACT  Late-bound via CreateObject/GetObject. bExcelStartedByUs prevents
'   quitting a pre-existing user session. All COM refs released on every exit path.
'
' ENTRY POINT  Call ExportLengthInRegion([ZoneLevel], [Filepath], [ExcelVisible])
'   Filepath empty → ARES_Zone_Export_Use_Dialog: True = Save-As dialog, False = auto path.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, ARESConstants, ErrorHandlerClass, FileDialogs, GetElements, LicenseManager

Option Explicit

Private Const HEADER_STYLE      As String = "Line Style"
Private Const HEADER_LEVEL      As String = "Level"
Private Const HEADER_COLOR      As String = "Color"
Private Const HEADER_LENGTH     As String = "Total Length (master units)"
Private Const XL_OPENXML_FORMAT As Long   = 51   ' xlOpenXMLWorkbook (.xlsx)

' ============================================================
'  PUBLIC ENTRY POINT
' ============================================================

' ExportLengthInRegion
' --------------------
' Aggregates the length of every graphical element inside any zone and
' writes a per-level summary to a new Excel workbook.
'
' Parameters (both optional):
'   ZoneLevel : level holding the zone elements. Defaults to
'               ARESConfig.ARES_ZONING_OUTPUT_LEVEL when omitted/empty.
'   Filepath  : when provided, saves the workbook to that path in .xlsx
'               format and overwrites any existing file silently. When
'               omitted, the workbook is left unsaved and visible.
Public Sub ExportLengthInRegion(Optional ByVal ZoneLevel As String = "", _
                                Optional ByVal Filepath As String = "", _
                                Optional ByVal ExcelVisible As Boolean = True)

    On Error GoTo ErrorHandler

    ' --- AC-1: License guard at entry point ---
    If Not LicenseManager.IsLicenseValid() Then
        ShowStatus "ARES: License not valid - ExportLengthInRegion disabled"
        Exit Sub
    End If

    ' --- AC-4: Config must be initialised ---
    If Not ARESConfig.IsInitialized Then
        ErrorHandler.HandleError "ARESConfig not initialized", 0, "", "ExportLengthInRegion.ExportLengthInRegion"
		Exit Sub
    End If

    ' --- AC-5: Active model must exist ---
    If Not Application.HasActiveModelReference Then
        ErrorHandler.HandleError "No active model reference", 0, "", "ExportLengthInRegion.ExportLengthInRegion"
		ShowStatus "ARES: ExportLengthInRegion - no active model reference"
        Exit Sub
    End If

    ' --- AC-2 / AC-3: resolve effective zone level ---
    If Len(ZoneLevel) = 0 Then ZoneLevel = ARESConfig.ARES_ZONING_OUTPUT_LEVEL.Value
    If Len(ZoneLevel) = 0 Then
        ErrorHandler.HandleError "Zone level is empty (config ARES_ZONING_OUTPUT_LEVEL not set)", 0, "", "ExportLengthInRegion.ExportLengthInRegion"
		ShowStatus "ARES: ExportLengthInRegion - zone level not configured"
        Exit Sub
    End If

    ' --- AC-7: zone level must exist ---
    If Not GetElements.IsValidLevelName(ZoneLevel) Then
        ErrorHandler.HandleError "Zone level not found in ActiveDesignFile.Levels: " & ZoneLevel, 0, "", "ExportLengthInRegion.ExportLengthInRegion"
		ShowStatus "ARES: ExportLengthInRegion - zone level not found: " & ZoneLevel
        Exit Sub
    End If

    ' --- Resolve filepath: dialog or auto-generated depending on ARES_ZONE_EXPORT_USE_DIALOG ---
    If Len(Filepath) = 0 Then
        If UCase(Trim(ARESConfig.ARES_ZONE_EXPORT_USE_DIALOG.Value)) = "TRUE" Then
            Filepath = FileDialogs.ShowSaveDialog( _
                           "Export Length in Region", _
                           "", _
                           BuildDefaultFilename(), _
                           DIALOG_FILTER_XLSX, "xlsx")
            If Len(Filepath) = 0 Then
                ShowStatus "ARES: ExportLengthInRegion - export cancelled"
                Exit Sub
            End If
        Else
            Filepath = BuildDefaultFilepath()
        End If
    End If

    ShowStatus "ARES: ExportLengthInRegion - collecting zones on level " & ZoneLevel

    ' --- T3: collect zone elements ---
    Dim zones() As Element
    If Not CollectZones(ZoneLevel, zones) Then
        ' AC-6: warning already logged inside CollectZones.
        ShowStatus "ARES: ExportLengthInRegion - no zones on level " & ZoneLevel
        Exit Sub
    End If

    ' --- T4: union bbox of all zones ---
    Dim oZoneRange As Range3d
    If Not ComputeZoneUnionRange(zones, oZoneRange) Then
        ShowStatus "ARES: ExportLengthInRegion - failed to compute zone bbox, aborting"
        Exit Sub
    End If

    ' --- T5: coarse-scan candidates (graphical, bbox overlap) ---
    Dim oee As ElementEnumerator
    Set oee = CollectCandidates(oZoneRange)

    ShowStatus "ARES: ExportLengthInRegion - scanning candidates"

    ' --- Resolve group-by mode ---
    Dim sGroupBy As String
    sGroupBy = Trim(ARESConfig.ARES_ZONE_EXPORT_GROUP_BY.Value)
    If sGroupBy <> "Level" And sGroupBy <> "Color" Then sGroupBy = "Style"

    ' --- T7: aggregate lengths ---
    Dim oGroups       As Object   ' Scripting.Dictionary
    Dim lElementCount As Long
    Set oGroups = CreateObject("Scripting.Dictionary")
    AggregateLengths oee, zones, ZoneLevel, oGroups, lElementCount, sGroupBy

    ' --- T8: export to Excel (always create the workbook, even when empty — AC-8) ---
    WriteToExcel oGroups, Filepath, ExcelVisible, sGroupBy

    ShowStatus "ARES: ExportLengthInRegion complete - " & CStr(lElementCount) & " elements, " & CStr(oGroups.Count) & " groups (" & sGroupBy & ")"
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.ExportLengthInRegion"
End Sub

' ============================================================
'  ZONE COLLECTION (T3) + BBOX (T4) + CANDIDATE SCAN (T5)
' ============================================================

' CollectZones
' Returns True when at least one zone element was found on the given level.
' Returns False and emits a WARNING (AC-6) when the level is empty.
Private Function CollectZones(ByVal ZoneLevel As String, ByRef outZones() As Element) As Boolean
    On Error GoTo ErrorHandler

    Dim ee As ElementEnumerator
    Set ee = GetElements.ByEE(Levels:=Array(ZoneLevel), _
                              ElTypes:=Array(msdElementTypeCellHeader, _
                                             msdElementTypeShape, _
                                             msdElementTypeComplexShape, _
                                             msdElementTypeEllipse))
    outZones = ee.BuildArrayFromContents

    If Not HasElements(outZones) Then
        ErrorHandler.HandleError "No zones found on level: " & ZoneLevel, 0, "", "ExportLengthInRegion.CollectZones"
		CollectZones = False
        Exit Function
    End If

    CollectZones = True
    Exit Function

ErrorHandler:
    CollectZones = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.CollectZones"
End Function

' ComputeZoneUnionRange
' Returns True and populates outRange with the world-space axis-aligned bbox
' containing every zone in the array. Returns False on any error, so callers
' never accidentally pass a zeroed range to IncludeOnlyWithinRange (which
' would silently yield zero candidates — see M-5 in the review).
Private Function ComputeZoneUnionRange(ByRef zones() As Element, _
                                       ByRef outRange As Range3d) As Boolean
    On Error GoTo ErrorHandler

    Dim oResult As Range3d
    Dim oCur    As Range3d
    Dim i       As Long

    ' Seed with the first zone's range, then accumulate by Range3dUnion.
    oResult = zones(LBound(zones)).Range
    For i = LBound(zones) + 1 To UBound(zones)
        oCur = zones(i).Range
        oResult = Range3dUnion(oResult, oCur)
    Next i

    outRange = oResult
    ComputeZoneUnionRange = True
    Exit Function

ErrorHandler:
    ComputeZoneUnionRange = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.ComputeZoneUnionRange"
End Function

' CollectCandidates
' Returns a lazy ElementEnumerator for all length-supported element types whose bbox
' overlaps the zone union range. Avoids materialising the full candidate set into memory.
Private Function CollectCandidates(ByRef oRange As Range3d) As ElementEnumerator
    On Error GoTo ErrorHandler

    Set CollectCandidates = GetElements.ByEE(Range:=oRange, _
                                              ElTypes:=Array(msdElementTypeLine, _
                                                             msdElementTypeArc, _
                                                             msdElementTypeLineString, _
                                                             msdElementTypeShape, _
                                                             msdElementTypeComplexString, _
                                                             msdElementTypeComplexShape))
    Exit Function

ErrorHandler:
    Set CollectCandidates = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.CollectCandidates"
End Function

' HasElements
' Safe bounds check (project pattern §6) — returns False for an uninitialised array.
Private Function HasElements(ByRef arr() As Element) As Boolean
    On Error Resume Next
    HasElements = False
    If UBound(arr) <> -1 Then HasElements = True
    On Error GoTo 0
End Function

' ============================================================
'  AGGREGATION (T7)
' ============================================================

' AggregateLengths
' Walks the candidate enumerator, skips zone-level elements (AC-9).
' Calls GetPartialLengthInsideZones directly — returns 0 when no overlap.
' MicroStation scans never return duplicate IDs, so no deduplication is needed.
Private Sub AggregateLengths(ByVal oee As ElementEnumerator, _
                              ByRef oZones() As Element, _
                              ByVal sZoneLevelName As String, _
                              ByRef oOutGroups As Object, _
                              ByRef lOutElementCount As Long, _
                              ByVal sGroupBy As String)
    On Error GoTo ErrorHandler

    Dim oEl  As Element
    Dim sKey As String
    Dim dLen As Double

    lOutElementCount = 0
    If oee Is Nothing Then Exit Sub

    Do While oee.MoveNext
        Set oEl = oee.Current
        If oEl.Level.Name <> sZoneLevelName Then
            dLen = Length.GetPartialLengthInsideZones(oEl, oZones)
            If dLen > 0 Then
                Select Case sGroupBy
                    Case "Level" : sKey = oEl.Level.Name
                    Case "Color" : sKey = CStr(oEl.Color)
                    Case Else    : sKey = oEl.LineStyle.Name
                End Select
                If oOutGroups.Exists(sKey) Then
                    oOutGroups(sKey) = oOutGroups(sKey) + dLen
                Else
                    oOutGroups.Add sKey, dLen
                End If
                lOutElementCount = lOutElementCount + 1
            End If
        End If
    Loop

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.AggregateLengths"
End Sub

' ============================================================
'  EXCEL EXPORT (T8)
' ============================================================

' Returns the directory where the export file will be saved.
' Uses the active design file's folder; falls back to the user's Desktop.
Private Function BuildDefaultDirectory() As String
    On Error Resume Next
    Dim sDir As String
    If Not ActiveDesignFile Is Nothing Then sDir = ActiveDesignFile.Path
    If Len(sDir) = 0 Then sDir = Environ("USERPROFILE") & "\Desktop"
    BuildDefaultDirectory = sDir
End Function

' Returns the default timestamped filename for the export workbook.
Private Function BuildDefaultFilename() As String
    BuildDefaultFilename = "ARES_LengthInRegion_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"
End Function

' Returns a full timestamped .xlsx path (directory + filename).
Private Function BuildDefaultFilepath() As String
    BuildDefaultFilepath = BuildDefaultDirectory() & "\" & BuildDefaultFilename()
End Function

' WriteToExcel
' Late-binds Excel, writes headers + one row per level (alphabetically
' sorted, case-insensitive), optionally saves to Filepath.
'
' COM lifecycle:
'   - bExcelStartedByUs tracks whether this sub created the Excel session.
'     Set to True OPTIMISTICALLY before CreateObject so a partially-created
'     Excel process (broken Office install, ActiveX disabled, OOM) is still
'     reachable for cleanup. Demoted to False only when we proved we reused
'     an existing session.
'   - Workbook close is ALWAYS attempted in the error path (regardless of
'     bExcelStartedByUs) so a phantom empty workbook is never left in the
'     user's pre-existing Excel session.
'   - Excel quit is gated on bExcelStartedByUs (AC-17: never quit the user's
'     pre-existing session).
'   - Cleanup: re-arms On Error Resume Next so a cleanup-time COM glitch
'     cannot escape and mask the root error in the caller's log.
Private Sub WriteToExcel(ByRef oLevels As Object, ByVal Filepath As String, _
                         ByVal bVisible As Boolean, ByVal sGroupBy As String)

    Dim xlApp             As Object
    Dim xlBook            As Object
    Dim xlSheet           As Object
    Dim bExcelStartedByUs As Boolean
    Dim sortedKeys()      As String
    Dim i                 As Long
    Dim sKey              As String
    Dim nRound            As Byte

    On Error GoTo ErrorHandler

    ' Resolve rounding precision from config; guard against reserved error value (255).
    Dim sRound As String
    sRound = ARESConfig.ARES_ZONE_EXPORT_ROUND.Value
    If Len(sRound) = 0 Then sRound = ARESConfig.ARES_ZONE_EXPORT_ROUND.defaultValue
    nRound = CByte(sRound)
    If nRound = ARES_RND_ERROR_VALUE Then nRound = CByte(ARESConfig.ARES_ZONE_EXPORT_ROUND.defaultValue)

    ' (1) Reuse existing Excel session if user already has one (AC-17),
    '     otherwise start our own and remember it for cleanup.
    '     bExcelStartedByUs defaults True so a half-initialised Excel process
    '     from a failed CreateObject is still quit in the error path (H-1).
    bExcelStartedByUs = True
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    On Error GoTo ErrorHandler
    If Not xlApp Is Nothing Then
        ' Successfully reused: not ours to quit on cleanup.
        bExcelStartedByUs = False
    Else
        Set xlApp = CreateObject("Excel.Application")
    End If

    ' (2) New workbook, name the first sheet.
    Set xlBook  = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    xlSheet.Name = ARESConfig.ARES_ZONE_EXPORT_SHEET_NAME.Value

    ' (3) Headers (AC-15).
    Dim sGroupHeader As String
    Select Case sGroupBy
        Case "Level" : sGroupHeader = HEADER_LEVEL
        Case "Color" : sGroupHeader = HEADER_COLOR
        Case Else    : sGroupHeader = HEADER_STYLE
    End Select
    xlSheet.Cells(1, 1).Value = sGroupHeader
    xlSheet.Cells(1, 2).Value = HEADER_LENGTH

    ' (4) Sort level names case-insensitively (AC-12) and write data rows.
    If oLevels.Count > 0 Then
        sortedKeys = SortedKeysCI(oLevels)
        For i = LBound(sortedKeys) To UBound(sortedKeys)
            sKey = sortedKeys(i)
            xlSheet.Cells(i - LBound(sortedKeys) + 2, 1).Value = sKey
            xlSheet.Cells(i - LBound(sortedKeys) + 2, 2).Value = Round(oLevels(sKey), nRound)
        Next i
    End If

    ' (5) Save when a path is provided (AC-18 / AC-19).
    If Len(Filepath) > 0 Then
        xlApp.DisplayAlerts = False
        xlBook.SaveAs Filepath, XL_OPENXML_FORMAT
        xlApp.DisplayAlerts = True
    End If

    ' (6) Surface the workbook or close the headless session.
    If bVisible Then
        xlApp.Visible = True
    Else
        ' Headless export: workbook is already saved; release the session silently.
        ' AC-17: only quit Excel when we started it — never kill the user's session.
        On Error Resume Next
        xlBook.Close False
        If bExcelStartedByUs Then xlApp.Quit
        On Error GoTo ErrorHandler
    End If

    On Error Resume Next
    Set xlSheet = Nothing
    Set xlBook  = Nothing
    Set xlApp   = Nothing
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.WriteToExcel"
	On Error Resume Next
    If Not xlApp Is Nothing Then xlApp.DisplayAlerts = True
    If Not xlBook Is Nothing Then xlBook.Close False
    If bExcelStartedByUs Then
        If Not xlApp Is Nothing Then xlApp.Quit
    End If
    Set xlSheet = Nothing
    Set xlBook  = Nothing
    Set xlApp   = Nothing
End Sub

' SortedKeysCI
' Returns a 0-based String() of the dictionary keys sorted case-insensitive.
' Bubble-sort is fine — level counts are small (< 100 in any realistic file).
Private Function SortedKeysCI(ByRef oDict As Object) As String()
    On Error GoTo ErrorHandler

    Dim keys() As String
    Dim i      As Long
    Dim j      As Long
    Dim tmp    As String
    Dim n      As Long
    Dim v      As Variant

    n = oDict.Count
    ReDim keys(0 To n - 1)
    i = 0
    For Each v In oDict.Keys
        keys(i) = CStr(v)
        i = i + 1
    Next v

    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If StrComp(keys(j), keys(j + 1), vbTextCompare) > 0 Then
                tmp = keys(j)
                keys(j) = keys(j + 1)
                keys(j + 1) = tmp
            End If
        Next j
    Next i

    SortedKeysCI = keys
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.SortedKeysCI"
End Function
