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
'   ARES_Zone_Export_Level (| -delimited) restricts measured candidates to those levels; empty = all levels.
'
' GROUPING  ARES_Zone_Export_Group_By ∈ {Style, Level, Color} — the group key of each measured
'   element. Unchanged; the classic path writes one global row per group (byte-identical).
' PER-ZONE SPLIT  ARES_Zone_Export_Per_Zone = True additionally splits every group by zone →
'   LONG FORMAT (Zone | <Style/Level/Color header> | Total Length), one row per (zone × group key).
'   Length is attributed PER ZONE (GetPartialLengthInsideZones called once per zone with a
'   1-element array). Each zone is labeled by the value of ARES_Zone_Export_Zone_Property read on
'   the ZONE element, else "Zone <n>" (empty var silent; a non-member name logs + one-shot status).
'   Same-value zones stay on SEPARATE rows (the aggregation key uses the zone scan index — Option A).
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, ARESConstants, ErrorHandlerClass, FileDialogs, GetElements, CustomPropertyHandler

Option Explicit

Private Const HEADER_STYLE      As String = "Line Style"
Private Const HEADER_LEVEL      As String = "Level"
Private Const HEADER_COLOR      As String = "Color"
Private Const HEADER_LENGTH     As String = "Total Length (master units)"
Private Const HEADER_ZONE       As String = "Zone"           ' per-zone split: column 1 header
Private Const HEADER_ID         As String = "ID"             ' group-by = ID (DLong): column header
Private Const KEY_SEP           As String = vbTab            ' composite key separator (zoneIndex & KEY_SEP & group-key)
Private Const XL_OPENXML_FORMAT As Long   = 51   ' xlOpenXMLWorkbook (.xlsx)
' Late-bound Excel enum values (named xl* constants are unavailable under CreateObject).
Private Const XL_EDGE_TOP        As Long   = 8      ' Borders() index: top edge
Private Const XL_EDGE_BOTTOM     As Long   = 9      ' Borders() index: bottom edge
Private Const XL_EDGE_LEFT       As Long   = 7      ' Borders() index: left edge
Private Const XL_EDGE_RIGHT      As Long   = 10     ' Borders() index: right edge
Private Const XL_INSIDE_VERTICAL As Long  = 11      ' Borders() index: inter-column lines
Private Const XL_LINE_CONTINUOUS As Long  = 1      ' xlContinuous
Private Const XL_WEIGHT_THIN     As Long   = 2      ' xlThin
Private Const XL_V_ALIGN_CENTER  As Long   = -4108 ' xlCenter (vertical alignment)

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

    ' --- AC-4: Config must be initialised ---
    If Not ARESConfig.IsInitialized Then
        ErrorHandler.HandleError "ARESConfig not initialized", 0, "", "ExportLengthInRegion.ExportLengthInRegion"
		Exit Sub
    End If

    ' --- AC-5: Active model must exist ---
    If Not Application.HasActiveModelReference Then
        ErrorHandler.HandleError "No active model reference", 0, "", "ExportLengthInRegion.ExportLengthInRegion"
		ShowStatusT "ZoneExportNoActiveModel"
        Exit Sub
    End If

    ' --- AC-2 / AC-3: resolve effective zone level ---
    If Len(ZoneLevel) = 0 Then ZoneLevel = ARESConfig.ARES_ZONING_OUTPUT_LEVEL.Value
    If Len(ZoneLevel) = 0 Then
        ErrorHandler.HandleError "Zone level is empty (config ARES_ZONING_OUTPUT_LEVEL not set)", 0, "", "ExportLengthInRegion.ExportLengthInRegion"
		ShowStatusT "ZoneExportLevelNotConfigured"
        Exit Sub
    End If

    ' --- AC-7: zone level must exist ---
    If Not GetElements.IsValidLevelName(ZoneLevel) Then
        ErrorHandler.HandleError "Zone level not found in ActiveDesignFile.Levels: " & ZoneLevel, 0, "", "ExportLengthInRegion.ExportLengthInRegion"
		ShowStatus GetTranslation("ZoneExportLevelNotFound", ZoneLevel)
        Exit Sub
    End If

    ' --- Resolve optional candidate level filter (ARES_Zone_Export_Level) ---
    '     Empty = all levels. Level names that do not exist in the file are ignored
    '     (non-fatal) and reported via a translated status; the export runs on the
    '     valid subset, or on all levels when none of the named levels exist.
    Dim sLevelFilter   As String
    Dim filterLevels() As String
    Dim sIgnoredLevels As String
    Dim nFilterLevels  As Long
    sLevelFilter = Trim(ARESConfig.ARES_ZONE_EXPORT_LEVEL.Value)
    nFilterLevels = ResolveFilterLevels(sLevelFilter, filterLevels, sIgnoredLevels)
    If Len(sIgnoredLevels) > 0 Then
        ErrorHandler.HandleError "Ignored non-existent export filter level(s): " & sIgnoredLevels, 0, "", "ExportLengthInRegion.ExportLengthInRegion"
        ShowStatus GetTranslation("ZoneExportFilterLevelsIgnored", sIgnoredLevels)
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
                ShowStatusT "ZoneExportCancelled"
                Exit Sub
            End If
        Else
            Filepath = BuildDefaultFilepath()
        End If
    End If

    ' --- T3: collect zone elements ---
    Dim zones() As Element
    If Not CollectZones(ZoneLevel, zones) Then
        ' AC-6: warning already logged inside CollectZones.
        ShowStatus GetTranslation("ZoneExportNoZones", ZoneLevel)
        Exit Sub
    End If

    ' --- T4: union bbox of all zones ---
    Dim oZoneRange As Range3d
    If Not ComputeZoneUnionRange(zones, oZoneRange) Then
        ErrorHandler.HandleError "failed to compute zone bbox, aborting", 0, "", "ExportLengthInRegion.ExportLengthInRegion": ShowStatusT "ZoneExportFailed"
        Exit Sub
    End If

    ' --- T5: coarse-scan candidates (graphical, bbox overlap, optional level filter) ---
    Dim oee As ElementEnumerator
    Set oee = CollectCandidates(oZoneRange, filterLevels, nFilterLevels)

    ' --- Resolve group-by key ---
    Dim sGroupBy As String
    sGroupBy = Trim(ARESConfig.ARES_ZONE_EXPORT_GROUP_BY.Value)
    If sGroupBy <> "Level" And sGroupBy <> "Color" And sGroupBy <> "ID" Then sGroupBy = "Style"

    ' --- Resolve per-zone split (independent axis from the group-by key) ---
    Dim bPerZone As Boolean
    bPerZone = (UCase(Trim(ARESConfig.ARES_ZONE_EXPORT_PER_ZONE.Value)) = "TRUE")

    ' --- T7: aggregate lengths (classic global, or additionally split per zone) ---
    Dim oGroups       As Object   ' Scripting.Dictionary
    Dim nElementCount As Long
    Dim zoneLabels()  As String
    Set oGroups = CreateObject("Scripting.Dictionary")
    If bPerZone Then
        Dim sLabelProp As String
        sLabelProp = ResolveZoneLabelProperty()
        BuildZoneLabels zones, sLabelProp, zoneLabels
        AggregateByZoneAndProperty oee, zones, zoneLabels, ZoneLevel, sGroupBy, oGroups, nElementCount
    Else
        AggregateLengths oee, zones, ZoneLevel, oGroups, nElementCount, sGroupBy
    End If

    ' --- T8: export to Excel (always create the workbook, even when empty — AC-8) ---
    WriteToExcel oGroups, Filepath, ExcelVisible, sGroupBy, bPerZone, zoneLabels

    If bPerZone Then
        ShowStatus GetTranslation("ZoneExportCompletePerZone", nElementCount, oGroups.Count, sGroupBy)
    Else
        ShowStatus GetTranslation("ZoneExportComplete", nElementCount, oGroups.Count, sGroupBy)
    End If
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
' When nFilterLevels > 0, the scan is additionally restricted to filterLevels (AND).
Private Function CollectCandidates(ByRef oRange As Range3d, _
                                   ByRef filterLevels() As String, _
                                   ByVal nFilterLevels As Long) As ElementEnumerator
    On Error GoTo ErrorHandler

    If nFilterLevels > 0 Then
        Set CollectCandidates = GetElements.ByEE(Levels:=filterLevels, _
                                                 Range:=oRange, _
                                                 ElTypes:=CandidateTypes())
    Else
        Set CollectCandidates = GetElements.ByEE(Range:=oRange, _
                                                 ElTypes:=CandidateTypes())
    End If
    Exit Function

ErrorHandler:
    Set CollectCandidates = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.CollectCandidates"
End Function

' CandidateTypes
' The length-supported element types scanned as export candidates.
Private Function CandidateTypes() As Variant
    CandidateTypes = Array(msdElementTypeLine, _
                           msdElementTypeArc, _
                           msdElementTypeLineString, _
                           msdElementTypeShape, _
                           msdElementTypeComplexString, _
                           msdElementTypeComplexShape)
End Function

' ResolveFilterLevels
' Parses the |-delimited ARES_Zone_Export_Level value into a 0-based array of
' trimmed, existing level names (empty tokens dropped). Non-existent names are
' NOT kept but accumulated into outIgnored (|-joined) so the caller can report
' them. Returns the count of valid (existing) names; 0 when the filter is empty
' or every named level is missing.
Private Function ResolveFilterLevels(ByVal sFilter As String, _
                                     ByRef outNames() As String, _
                                     ByRef outIgnored As String) As Long
    On Error GoTo ErrorHandler

    Dim parts() As String
    Dim sName   As String
    Dim i       As Long
    Dim n       As Long

    n = 0
    outIgnored = ""
    If Len(Trim(sFilter)) = 0 Then
        ResolveFilterLevels = 0
        Exit Function
    End If

    parts = Split(sFilter, ARES_VAR_DELIMITER)
    ReDim outNames(0 To UBound(parts) - LBound(parts))
    For i = LBound(parts) To UBound(parts)
        sName = Trim(parts(i))
        If Len(sName) > 0 Then
            If GetElements.IsValidLevelName(sName) Then
                outNames(n) = sName
                n = n + 1
            Else
                If Len(outIgnored) > 0 Then outIgnored = outIgnored & ARES_VAR_DELIMITER
                outIgnored = outIgnored & sName
            End If
        End If
    Next i

    If n > 0 Then ReDim Preserve outNames(0 To n - 1)
    ResolveFilterLevels = n
    Exit Function

ErrorHandler:
    ResolveFilterLevels = 0
    outIgnored = ""
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.ResolveFilterLevels"
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
                              ByRef nOutElementCount As Long, _
                              ByVal sGroupBy As String)
    On Error GoTo ErrorHandler

    Dim oEl  As Element
    Dim sKey As String
    Dim dLen As Double

    nOutElementCount = 0
    If oee Is Nothing Then Exit Sub

    Do While oee.MoveNext
        Set oEl = oee.Current
        If oEl.Level.Name <> sZoneLevelName Then
            dLen = Length.GetPartialLengthInsideZones(oEl, oZones)
            If dLen > 0 Then
                Select Case sGroupBy
                    Case "Level" : sKey = oEl.Level.Name
                    Case "Color" : sKey = CStr(oEl.Color)
                    Case "ID"    : sKey = DLongToString(oEl.ID)
                    Case Else    : sKey = oEl.LineStyle.Name
                End Select
                If oOutGroups.Exists(sKey) Then
                    oOutGroups(sKey) = oOutGroups(sKey) + dLen
                Else
                    oOutGroups.Add sKey, dLen
                End If
                nOutElementCount = nOutElementCount + 1
            End If
        End If
    Loop

    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.AggregateLengths"
End Sub

' ============================================================
'  PER-ZONE SPLIT (aggregate each group additionally by zone)
' ============================================================

' NameInList
' Case-insensitive membership test over a 0-based names array (GetCustomPropertyNames output).
' Split() always returns an allocated array, so LBound/UBound are safe here.
Private Function NameInList(ByVal sName As String, ByRef names() As String) As Boolean
    On Error GoTo ErrorHandler

    NameInList = False
    Dim i As Long
    For i = LBound(names) To UBound(names)
        If StrComp(Trim(names(i)), sName, vbTextCompare) = 0 Then
            NameInList = True
            Exit Function
        End If
    Next i
    Exit Function

ErrorHandler:
    NameInList = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.NameInList"
End Function

' ResolveZoneLabelProperty
' Resolves ARES_ZONE_EXPORT_ZONE_PROPERTY once (per-zone split only):
'   empty                          → "" (silent; zones fall back to "Zone <n>")
'   non-empty, member of the list  → the name (zones labeled by its value on each zone element)
'   non-empty, NOT a member        → log (English) + one-shot ZoneExportZonePropertyInvalid
'                                    status + "" (zones fall back to "Zone <n>"). Never aborts.
Private Function ResolveZoneLabelProperty() As String
    On Error GoTo ErrorHandler

    ResolveZoneLabelProperty = ""
    Dim sName   As String
    Dim names() As String
    sName = Trim(ARESConfig.ARES_ZONE_EXPORT_ZONE_PROPERTY.Value)
    If Len(sName) = 0 Then Exit Function          ' unconfigured → silent "Zone <n>"

    names = CustomPropertyHandler.GetCustomPropertyNames()
    If NameInList(sName, names) Then
        ResolveZoneLabelProperty = sName
        Exit Function
    End If

    ErrorHandler.HandleError "Zone property set but not a member of ARES_Custom_Property_List: '" & sName & "'", 0, "", "ExportLengthInRegion.ResolveZoneLabelProperty"
    ShowStatusT "ZoneExportZonePropertyInvalid"          ' non-fatal — zones fall back to "Zone <n>"
    Exit Function

ErrorHandler:
    ResolveZoneLabelProperty = ""
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.ResolveZoneLabelProperty"
End Function

' BuildZoneLabels
' Fills outLabels (0-based, indexed by zone scan position z = i - LBound(zones)) with each
' zone's display label via ResolveZoneLabel, using the chosen zone-property (sLabelProp;
' "" ⇒ every zone gets "Zone <n>"). The 0-based indexing matches the aggregation key's zone index.
Private Sub BuildZoneLabels(ByRef zones() As Element, ByVal sLabelProp As String, _
                            ByRef outLabels() As String)
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim z As Long
    ReDim outLabels(0 To UBound(zones) - LBound(zones))
    For i = LBound(zones) To UBound(zones)
        z = i - LBound(zones)
        outLabels(z) = ResolveZoneLabel(zones(i), z, sLabelProp)
    Next i
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.BuildZoneLabels"
End Sub

' ResolveZoneLabel
' Returns the zone element's label: the value of sLabelProp read on the zone element
' (GetPropertyValueFromElement, name = ItemType = property), else "Zone <idx+1>".
' Falls back to the positional label when sLabelProp = "" or the zone has no/empty value.
' idx is 0-based (scan order); the fallback is 1-based for human display.
Private Function ResolveZoneLabel(ByVal oZone As Element, ByVal idx As Long, _
                                  ByVal sLabelProp As String) As String
    On Error GoTo ErrorHandler

    ResolveZoneLabel = "Zone " & (idx + 1)
    If Len(sLabelProp) = 0 Then Exit Function

    Dim vVal As Variant
    vVal = CustomPropertyHandler.GetPropertyValueFromElement(oZone, sLabelProp, sLabelProp)
    If Not IsNull(vVal) Then
        Dim sVal As String
        sVal = Trim(CStr(vVal))
        If Len(sVal) > 0 Then ResolveZoneLabel = sVal
    End If
    Exit Function

ErrorHandler:
    ' Keep the positional "Zone <n>" fallback already assigned above.
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.ResolveZoneLabel"
End Function

' AggregateByZoneAndProperty
' Per-zone aggregation. Mirrors AggregateLengths but attributes length PER ZONE: for each
' candidate, loops the zones and calls GetPartialLengthInsideZones with a 1-element array (the
' tested engine, UNCHANGED) to get the element's length inside THAT zone. The group key is the
' SAME as AggregateLengths (Style→LineStyle.Name, Level→Level.Name, Color→CStr(Color)). Composite
' key = Format(z, "0000") & KEY_SEP & <group key>, where z is the 0-based zone scan index — so two
' zones sharing a label stay on SEPARATE rows (Option A) and SortedKeysCI orders zone-major
' (zero-padded numeric), then group key. zoneLabels is display-only (consumed by WriteToExcel via
' the same z index); the aggregation itself keys on the index, not the label. nOutElementCount
' counts candidates contributing length to >= 1 zone.
Private Sub AggregateByZoneAndProperty(ByVal oee As ElementEnumerator, _
                                       ByRef oZones() As Element, _
                                       ByRef zoneLabels() As String, _
                                       ByVal sZoneLevelName As String, _
                                       ByVal sGroupBy As String, _
                                       ByRef oOutGroups As Object, _
                                       ByRef nOutElementCount As Long)
    On Error GoTo ErrorHandler

    Dim oEl              As Element
    Dim oneZone(0 To 0)  As Element
    Dim i                As Long
    Dim z                As Long
    Dim dLen             As Double
    Dim sGbKey           As String
    Dim sKey             As String
    Dim bCounted         As Boolean

    nOutElementCount = 0
    If oee Is Nothing Then Exit Sub

    Do While oee.MoveNext
        Set oEl = oee.Current
        If oEl.Level.Name <> sZoneLevelName Then
            Select Case sGroupBy
                Case "Level" : sGbKey = oEl.Level.Name
                Case "Color" : sGbKey = CStr(oEl.Color)
                Case "ID"    : sGbKey = DLongToString(oEl.ID)
                Case Else    : sGbKey = oEl.LineStyle.Name
            End Select
            bCounted = False
            For i = LBound(oZones) To UBound(oZones)
                z = i - LBound(oZones)
                Set oneZone(0) = oZones(i)
                dLen = Length.GetPartialLengthInsideZones(oEl, oneZone)
                If dLen > 0 Then
                    sKey = Format(z, "0000") & KEY_SEP & sGbKey
                    If oOutGroups.Exists(sKey) Then
                        oOutGroups(sKey) = oOutGroups(sKey) + dLen
                    Else
                        oOutGroups.Add sKey, dLen
                    End If
                    bCounted = True
                End If
            Next i
            If bCounted Then nOutElementCount = nOutElementCount + 1
        End If
    Loop
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.AggregateByZoneAndProperty"
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
' In per-zone mode (bLongFormat = True), writes 3 columns (Zone | <group-by header> | Total
' Length), splitting each composite key (zoneIndex & KEY_SEP & group key) — column 1 is the zone's
' display label from zoneLabels(zoneIndex). Otherwise the classic 2-column path (untouched,
' byte-identical). zoneLabels is only used for the long-format column 1.
' Long-format rows are then grouped visually (MergeZoneBlocks): each zone's label cell is merged
' over its rows and a thin border separates consecutive zones.
Private Sub WriteToExcel(ByRef oLevels As Object, ByVal Filepath As String, _
                         ByVal bVisible As Boolean, ByVal sGroupBy As String, _
                         ByVal bLongFormat As Boolean, ByRef zoneLabels() As String)

    Dim xlApp             As Object
    Dim xlBook            As Object
    Dim xlSheet           As Object
    Dim bExcelStartedByUs As Boolean
    Dim sortedKeys()      As String
    Dim i                 As Long
    Dim z                 As Long
    Dim sKey              As String
    Dim nRound            As Byte
    Dim sGroupHeader      As String
    Dim keyParts()        As String

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

    ' (3) Headers (AC-15). Per-zone mode writes 3 columns; classic writes 2 (untouched).
    '     Column 2's header is the group-by header in BOTH modes.
    Select Case sGroupBy
        Case "Level" : sGroupHeader = HEADER_LEVEL
        Case "Color" : sGroupHeader = HEADER_COLOR
        Case "ID"    : sGroupHeader = HEADER_ID
        Case Else    : sGroupHeader = HEADER_STYLE
    End Select
    If bLongFormat Then
        xlSheet.Cells(1, 1).Value = HEADER_ZONE
        xlSheet.Cells(1, 2).Value = sGroupHeader
        xlSheet.Cells(1, 3).Value = HEADER_LENGTH
    Else
        xlSheet.Cells(1, 1).Value = sGroupHeader
        xlSheet.Cells(1, 2).Value = HEADER_LENGTH
    End If

    ' (4) Sort keys case-insensitively (AC-12) and write data rows.
    '     Classic: (group key, length). Per-zone: composite key (zoneIndex & KEY_SEP & group key)
    '     → zone label (col 1, via zoneLabels(zoneIndex)) + group key (col 2), length in col 3.
    If oLevels.Count > 0 Then
        sortedKeys = SortedKeysCI(oLevels)
        For i = LBound(sortedKeys) To UBound(sortedKeys)
            sKey = sortedKeys(i)
            If bLongFormat Then
                keyParts = Split(sKey, KEY_SEP)
                z = CLng(keyParts(0))
                xlSheet.Cells(i - LBound(sortedKeys) + 2, 1).Value = zoneLabels(z)
                If UBound(keyParts) >= 1 Then _
                    xlSheet.Cells(i - LBound(sortedKeys) + 2, 2).Value = keyParts(1)
                xlSheet.Cells(i - LBound(sortedKeys) + 2, 3).Value = Round(oLevels(sKey), nRound)
            Else
                xlSheet.Cells(i - LBound(sortedKeys) + 2, 1).Value = sKey
                xlSheet.Cells(i - LBound(sortedKeys) + 2, 2).Value = Round(oLevels(sKey), nRound)
            End If
        Next i
    End If

    ' (4b) Per-zone visual grouping (long format only): merge each zone's label cell over its rows
    '      and draw a separator border between zones, so a multi-level zone reads as one block and
    '      consecutive same-named zones stay distinct (runs keyed by zone index, not label).
    If bLongFormat And oLevels.Count > 0 Then MergeZoneBlocks xlSheet, xlApp, sortedKeys

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

' MergeZoneBlocks (per-zone long format only)
' Merges the Zone-label cell (column 1) across each contiguous run of rows sharing the same zone
' SCAN INDEX, vertically centres it, and draws a thin top border at each zone boundary. Rows are
' already zone-index-major and contiguous (composite key = zeroPaddedIndex & KEY_SEP & group key),
' so a run break = a change of the leading index. Keying on the index (not the label) keeps two
' consecutive zones with the SAME name in separate blocks - the whole point of the grouping.
Private Sub MergeZoneBlocks(ByVal xlSheet As Object, ByVal xlApp As Object, ByRef sortedKeys() As String)
    On Error GoTo ErrorHandler

    Dim n As Long
    n = UBound(sortedKeys) - LBound(sortedKeys) + 1
    If n <= 0 Then Exit Sub

    Dim bPrevAlerts As Boolean
    bPrevAlerts = xlApp.DisplayAlerts
    xlApp.DisplayAlerts = False          ' identical labels never prompt, but stay safe

    Dim i        As Long
    Dim r        As Long
    Dim zCur     As Long
    Dim zPrev    As Long
    Dim runStart As Long

    runStart = 2                         ' first data row (row 1 = header)
    zPrev = ZoneIndexOfKey(sortedKeys(LBound(sortedKeys)))
    For i = LBound(sortedKeys) + 1 To UBound(sortedKeys)
        r = i - LBound(sortedKeys) + 2
        zCur = ZoneIndexOfKey(sortedKeys(i))
        If zCur <> zPrev Then
            FinishZoneBlock xlSheet, runStart, r - 1
            runStart = r
            zPrev = zCur
        End If
    Next i
    FinishZoneBlock xlSheet, runStart, n + 1   ' last data row = 2 + n - 1
    FrameTable xlSheet, n + 1                   ' outer + inter-column borders, header & bottom lines

    xlApp.DisplayAlerts = bPrevAlerts
    Exit Sub

ErrorHandler:
    On Error Resume Next
    xlApp.DisplayAlerts = True
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.MergeZoneBlocks"
End Sub

' Merge + vertically centre the zone label over [rStart, rEnd]; for every block after the first,
' draw a thin top border across the 3 columns to separate it from the zone above.
Private Sub FinishZoneBlock(ByVal xlSheet As Object, ByVal rStart As Long, ByVal rEnd As Long)
    On Error GoTo ErrorHandler
    If rEnd > rStart Then
        With xlSheet.Range(xlSheet.Cells(rStart, 1), xlSheet.Cells(rEnd, 1))
            .Merge
            .VerticalAlignment = XL_V_ALIGN_CENTER
        End With
    End If
    If rStart > 2 Then
        SetBorder xlSheet.Range(xlSheet.Cells(rStart, 1), xlSheet.Cells(rStart, 3)).Borders(XL_EDGE_TOP)
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.FinishZoneBlock"
End Sub

' FrameTable (per-zone long format only)
' Frames the whole table: outer left/right edges + inter-column vertical lines over every row
' (header + data), a line under the header, and a line closing the bottom. The between-zone
' horizontal separators are drawn by FinishZoneBlock; here we add only the outer frame + verticals.
Private Sub FrameTable(ByVal xlSheet As Object, ByVal lastRow As Long)
    On Error GoTo ErrorHandler
    With xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(lastRow, 3))
        SetBorder .Borders(XL_EDGE_LEFT)
        SetBorder .Borders(XL_EDGE_RIGHT)
        SetBorder .Borders(XL_INSIDE_VERTICAL)
    End With
    SetBorder xlSheet.Range(xlSheet.Cells(1, 1), xlSheet.Cells(1, 3)).Borders(XL_EDGE_BOTTOM)
    SetBorder xlSheet.Range(xlSheet.Cells(lastRow, 1), xlSheet.Cells(lastRow, 3)).Borders(XL_EDGE_BOTTOM)
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.FrameTable"
End Sub

' Apply a thin continuous line to a single Border object (late-bound Excel).
Private Sub SetBorder(ByVal oBorder As Object)
    On Error GoTo ErrorHandler
    oBorder.LineStyle = XL_LINE_CONTINUOUS
    oBorder.Weight = XL_WEIGHT_THIN
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.SetBorder"
End Sub

' Leading zone scan index of a composite sort key (zeroPaddedIndex & KEY_SEP & group key).
Private Function ZoneIndexOfKey(ByVal sKey As String) As Long
    On Error GoTo ErrorHandler
    Dim parts() As String
    parts = Split(sKey, KEY_SEP)
    ZoneIndexOfKey = CLng(parts(0))
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInRegion.ZoneIndexOfKey"
End Function

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
