' CableReport.bas
' Description: For each HTAS cable (a Line element on a configured level), builds one Excel row
' summarizing its end-cell markers (Repere-family property), its linked text's Nature/Longueur
' (custom properties), and its trenching length broken down by soil-type zone (Coupe_Type-family
' property), pivoted into one column per distinct soil-type value and summed when several zones
' along the same cable share the same value.
'
' CABLE IDENTIFICATION  any element on ARES_CableReport_Cable_Level of a type supported by
'                        Length.GetPartialLengthInsideZones (Line/Arc/LineString/Shape/
'                        ComplexString/ComplexShape - same candidate set as ExportLengthInRegion).
' REPERE                the configured property (default "Repere") read off the nearest
'                        CellHeader within ARES_CableReport_Search_Radius of each endpoint
'                        (start/end, geometric proximity - NOT graphic group). Only the 4 OPEN
'                        types have two distinct ends (Line/Arc/LineString/ComplexString); a
'                        closed Shape/ComplexShape gets no Repere (still counted for length).
' NATURE / LONGUEUR     the configured properties (default "Nature"/"Longueur") read off whichever
'                        graphic-group member(s) carry them (Link.GetLink, NOT type-filtered - a
'                        Cell or any other grouped element qualifies, not just a Text), verbatim -
'                        no parsing/reformatting. Each property is resolved independently, so
'                        Nature and Longueur may even live on two different group members.
'                        Longueur is written to Excel as a real NUMBER when it is purely numeric
'                        (see TryAsNumber) - the content is still the property's, untouched, but a
'                        String "493,6" assigned to .Value would be stored as TEXT (Excel's object
'                        model parses an assigned string with the INVARIANT locale, not the user's),
'                        which is what made some cells text and others numeric.
' SOIL-TYPE BREAKDOWN   for each zone (Shape/ComplexShape on ARES_CableReport_Zone_Level) whose
'                        partial length inside it (Length.GetPartialLengthInsideZones) is > 0, the
'                        length is summed into a column keyed by that zone's configured property
'                        value (default "Coupe_Type"). Two zones sharing the same value for the
'                        same cable SUM into one column. Column order is alphabetical (case-
'                        insensitive), not scan order - deterministic run to run.
' SHARED TRENCH         a zone with >= 2 cables inside it (len > 0) is a shared trench, and it gets its
'                        OWN pivot column: the column key is "<label> (n)" ("CH2C (2)") instead of the
'                        plain "<label>", so a shared stretch reads as a column of its own next to the
'                        ordinary one rather than silently inflating it. The column is keyed on the
'                        SHARING GROUP, not just the count: two TA1 zones each shared by 2 cables but by
'                        different pairs are two trenches, hence two columns ("TA1 (2)", "TA1 (2) #2");
'                        two zones with the SAME members merge (two stretches of one shared trench add
'                        up). Several can coexist ("CH2C", "CH2C (2)", "CH2C (3)"), and one cable can
'                        appear in several - it may share its start with one cable and its end with
'                        another. Per-cable rows are NOT touched
'                        (each keeps its own length in that zone - cable-side truth); a 4th fixed column
'                        "Shared with" cross-lists the OTHER cables (Repere, else "#<Excel row>"), and a
'                        two-row FOOTER gives, per column, "Cable totals" (the rows added up as they
'                        are) and "Trench to dig (deduplicated)" = sum over zones of the MAX in-zone
'                        cable length - a zone is one trench segment, dug once, along its longest
'                        occupant. ASSUMPTION: one common zone per shared trench (RunOutline/MergeRegion/
'                        SplitRegion output); two DISTINCT zones overlapping each other are not detected,
'                        and one zone deliberately spanning two side-by-side trenches would be
'                        under-counted - both out of the confirmed use. Zone labels are read ONCE
'                        (BuildZoneLabels); a blank label contributes to nothing.
'
' EXCEL COM CONTRACT  Late-bound via CreateObject/GetObject. bExcelStartedByUs prevents quitting a
'   pre-existing user session. All COM refs released on every exit path. (Mirrors ExportLengthInRegion.)
'
' ENTRY POINT  Call CableReport([CableLevel], [ZoneLevel], [Filepath], [ExcelVisible])
'   ZoneLevel empty -> ARES_Outline_Output_Level. Filepath empty -> Save-As dialog (cancel aborts).
'   A missing linked text / end-cell is an expected drawing gap: the cell is left blank and the
'   row counted incomplete, never logged - see the final status line.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, ARESConstants, ErrorHandlerClass, FileDialogs, GetElements, CustomPropertyHandler, Link, Length

Option Explicit

Private Const KEY_SEP           As String = vbTab   ' composite key separator (cableKey & KEY_SEP & zone label)
Private Const XL_OPENXML_FORMAT As Long   = 51       ' xlOpenXMLWorkbook (.xlsx)

' ============================================================
'  PUBLIC ENTRY POINT
' ============================================================

Public Sub CableReport(Optional ByVal CableLevel As String = "", _
                       Optional ByVal ZoneLevel As String = "", _
                       Optional ByVal Filepath As String = "", _
                       Optional ByVal ExcelVisible As Boolean = True)

    On Error GoTo ErrorHandler

    If Not ARESConfig.IsInitialized Then
        ErrorHandler.HandleError "ARESConfig not initialized", 0, "", "CableReport.CableReport"
        Exit Sub
    End If

    If Not Application.HasActiveModelReference Then
        ErrorHandler.HandleError "No active model reference", 0, "", "CableReport.CableReport"
        ShowStatusT "CableReportNoActiveModel"
        Exit Sub
    End If

    ' --- Resolve cable level(s): |-delimited list (ARES_VAR_DELIMITER) - several cable-laying
    ' levels (e.g. BT/BTS/HTAS) do not all necessarily exist in a given drawing. Mirrors
    ' ExportLengthInRegion's candidate-level filter: each EXISTING name is kept, a missing one is
    ' ignored (non-fatal, reported), only an ALL-missing list aborts. ---
    If Len(CableLevel) = 0 Then CableLevel = ARESConfig.ARES_CABLEREPORT_CABLE_LEVEL.Value
    If Len(CableLevel) = 0 Then
        ErrorHandler.HandleError "Cable level is empty (config ARES_CableReport_Cable_Level not set)", 0, "", "CableReport.CableReport"
        ShowStatusT "CableReportLevelNotConfigured"
        Exit Sub
    End If

    Dim cableLevels()   As String
    Dim sIgnoredLevels  As String
    Dim nCableLevels    As Long
    nCableLevels = ResolveCableLevels(CableLevel, cableLevels, sIgnoredLevels)
    If nCableLevels = 0 Then
        ShowStatus GetTranslation("CableReportLevelNotFound", CableLevel)
        Exit Sub
    End If
    If Len(sIgnoredLevels) > 0 Then
        ' Status-bar only, not logged: a typo'd/renamed level in the list is a user config issue,
        ' not a fault - the export still runs fine on the valid subset.
        ShowStatus GetTranslation("CableReportLevelsIgnored", sIgnoredLevels)
    End If

    ' --- Resolve zone level (optional - a missing/invalid level degrades to a 3-column export) ---
    If Len(ZoneLevel) = 0 Then ZoneLevel = ARESConfig.ARES_CABLEREPORT_ZONE_LEVEL.Value
    If Len(ZoneLevel) = 0 Then ZoneLevel = ARESConfig.ARES_OUTLINE_OUTPUT_LEVEL.Value

    ' --- Resolve the 4 configured property names, each independently validated ---
    Dim sRepereProp   As String
    Dim sNatureProp   As String
    Dim sLongueurProp As String
    Dim sZoneProp     As String
    sRepereProp = ResolveConfiguredProperty(ARESConfig.ARES_CABLEREPORT_REPERE_PROPERTY.Value, "Repere")
    sNatureProp = ResolveConfiguredProperty(ARESConfig.ARES_CABLEREPORT_NATURE_PROPERTY.Value, "Nature")
    sLongueurProp = ResolveConfiguredProperty(ARESConfig.ARES_CABLEREPORT_LONGUEUR_PROPERTY.Value, "Longueur")
    sZoneProp = ResolveConfiguredProperty(ARESConfig.ARES_CABLEREPORT_ZONE_PROPERTY.Value, "Coupe_Type")

    ' --- Resolve output filepath (Save-As dialog; cancel aborts) ---
    If Len(Filepath) = 0 Then
        Filepath = FileDialogs.ShowSaveDialog( _
                       "Cable Report", _
                       "", _
                       BuildDefaultFilename(), _
                       DIALOG_FILTER_XLSX, "xlsx")
        If Len(Filepath) = 0 Then
            ShowStatusT "CableReportCancelled"
            Exit Sub
        End If
    End If

    ' --- Collect cables ---
    Dim cables() As Element
    If Not CollectCables(cableLevels, cables) Then
        ShowStatusT "CableReportNoCables"
        Exit Sub
    End If

    ' --- Collect zones (optional) ---
    Dim zones()   As Element
    Dim bHasZones As Boolean
    bHasZones = CollectZones(ZoneLevel, zones)
    If Not bHasZones Then ShowStatusT "CableReportNoZones"

    ' --- Zone labels resolved ONCE (not per cable x zone): each zone's soil-type value, "" when the
    ' property is absent/blank - a blank-labelled zone contributes to nothing (pivot, shared detection,
    ' footers). nZones = 0 when there are no zones or no zone property to read. ---
    Dim zoneLabels() As String
    Dim nZones       As Long
    nZones = 0
    If bHasZones And Len(sZoneProp) > 0 Then nZones = BuildZoneLabels(zones, sZoneProp, zoneLabels)

    ' --- Accumulate one row per cable + the raw per-(zone, cable) lengths ---
    ' The pivot COLUMNS cannot be decided here: a column key is "<label>" for a zone one cable runs
    ' through and "<label> (n)" for a shared trench, and n is only known once every cable has been
    ' measured against every zone. So this loop records raw lengths only; BuildPivotAndSharedTrenches
    ' derives the columns, the pivot, the cross-marks and the footers from them in one post-pass.
    Dim oRowData   As Object   ' Scripting.Dictionary: cableKey -> Array(sRepere, vNature, vLongueur)
    Dim oColumns   As Object   ' Scripting.Dictionary: column key -> True (presence only; sorted at write time)
    Dim oPivot     As Object   ' Scripting.Dictionary: cableKey & KEY_SEP & column key -> Double
    Dim oZoneCable As Object   ' Scripting.Dictionary: zoneIdx & KEY_SEP & cableKey -> Double (post-pass input)
    Set oRowData = CreateObject("Scripting.Dictionary")
    Set oColumns = CreateObject("Scripting.Dictionary")
    Set oPivot = CreateObject("Scripting.Dictionary")
    Set oZoneCable = CreateObject("Scripting.Dictionary")

    Dim rowOrder() As String
    ReDim rowOrder(0 To UBound(cables) - LBound(cables))

    Dim dRadius As Double
    dRadius = Val(ARESConfig.ARES_CABLEREPORT_SEARCH_RADIUS.Value)
    If dRadius <= 0 Then dRadius = Val(ARESConfig.ARES_CABLEREPORT_SEARCH_RADIUS.DefaultValue)

    Dim i             As Long
    Dim oEl           As Element
    Dim sCableKey     As String
    Dim sRepere       As String
    Dim vNature       As Variant
    Dim vLongueur     As Variant
    Dim bRowIncomplete As Boolean
    Dim nRowCount     As Long
    Dim nIncomplete   As Long

    For i = LBound(cables) To UBound(cables)
        Set oEl = cables(i)
        sCableKey = DLongToString(oEl.id)
        bRowIncomplete = False

        sRepere = ResolveCableRepere(oEl, sRepereProp, dRadius, bRowIncomplete)
        ResolveCableText oEl, sNatureProp, sLongueurProp, vNature, vLongueur, bRowIncomplete

        oRowData.Add sCableKey, Array(sRepere, vNature, vLongueur)
        rowOrder(i - LBound(cables)) = sCableKey
        nRowCount = nRowCount + 1
        If bRowIncomplete Then nIncomplete = nIncomplete + 1

        If nZones > 0 Then
            RecordZoneLengths oEl, sCableKey, zones, zoneLabels, oZoneCable
        End If
    Next i

    ' --- Post-pass: columns + pivot + shared-trench cross-marks + footers ---
    Dim oShared      As Object   ' cableKey -> ", "-joined Reperes of the OTHER cables sharing >= 1 zone
    Dim oCableTotal  As Object   ' column key -> plain sum of every cable's length (what the rows add up to)
    Dim oTrench      As Object   ' column key -> sum over zones of the MAX in-zone cable length (dug once)
    Dim nSharedZones As Long
    Set oShared = CreateObject("Scripting.Dictionary")
    Set oCableTotal = CreateObject("Scripting.Dictionary")
    Set oTrench = CreateObject("Scripting.Dictionary")
    nSharedZones = BuildPivotAndSharedTrenches(oZoneCable, zoneLabels, nZones, rowOrder, oRowData, _
                                               oColumns, oPivot, oShared, oCableTotal, oTrench)

    ' --- Write to Excel (always create the workbook, even with zero pivot columns) ---
    WriteToExcel oRowData, oColumns, oPivot, oShared, oCableTotal, oTrench, rowOrder, Filepath, ExcelVisible

    ShowStatus GetTranslation("CableReportComplete", nRowCount, oColumns.Count, nIncomplete, nSharedZones)
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport.CableReport"
End Sub

' ============================================================
'  CABLE / ZONE COLLECTION
' ============================================================

Private Function CollectCables(ByRef CableLevels() As String, ByRef outCables() As Element) As Boolean
    On Error GoTo ErrorHandler

    Dim ee As ElementEnumerator
    Set ee = GetElements.ByEE(Levels:=CableLevels, ElTypes:=CandidateTypes())
    outCables = ee.BuildArrayFromContents

    CollectCables = HasElements(outCables)
    Exit Function

ErrorHandler:
    CollectCables = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport.CollectCables"
End Function

' CandidateTypes
' The same length-supported element types as Length.GetPartialLengthInsideZones (mirrors
' ExportLengthInRegion.CandidateTypes) - a cable may be a straight Line, an Arc, a multi-vertex
' LineString, or a ComplexString/ComplexShape chain following a route, not just a 2-point Line.
Private Function CandidateTypes() As Variant
    CandidateTypes = Array(msdElementTypeLine, _
                           msdElementTypeArc, _
                           msdElementTypeLineString, _
                           msdElementTypeShape, _
                           msdElementTypeComplexString, _
                           msdElementTypeComplexShape)
End Function

' ResolveCableLevels
' Parses the |-delimited ARES_CableReport_Cable_Level value into a 0-based array of trimmed,
' EXISTING level names (empty tokens dropped). Non-existent names are NOT kept but accumulated
' into outIgnored (|-joined) so the caller can report them. Returns the count of valid (existing)
' names; 0 when every named level is missing. Mirrors ExportLengthInRegion.ResolveFilterLevels.
Private Function ResolveCableLevels(ByVal sLevels As String, ByRef outNames() As String, _
                                    ByRef outIgnored As String) As Long
    On Error GoTo ErrorHandler

    Dim parts() As String
    Dim sName   As String
    Dim i       As Long
    Dim n       As Long

    n = 0
    outIgnored = ""
    If Len(Trim(sLevels)) = 0 Then
        ResolveCableLevels = 0
        Exit Function
    End If

    parts = Split(sLevels, ARES_VAR_DELIMITER)
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
    ResolveCableLevels = n
    Exit Function

ErrorHandler:
    ResolveCableLevels = 0
    outIgnored = ""
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport.ResolveCableLevels"
End Function

' Returns True when at least one zone element was found on a configured, existing ZoneLevel.
' An empty ZoneLevel (nothing configured, no Outline fallback either) short-circuits to False
' without a scan attempt - the caller degrades to a 3-column export, never an abort.
Private Function CollectZones(ByVal ZoneLevel As String, ByRef outZones() As Element) As Boolean
    On Error GoTo ErrorHandler

    If Len(Trim(ZoneLevel)) = 0 Then
        CollectZones = False
        Exit Function
    End If

    Dim ee As ElementEnumerator
    Set ee = GetElements.ByEE(Levels:=Array(ZoneLevel), ElTypes:=Array(msdElementTypeShape, msdElementTypeComplexShape))
    outZones = ee.BuildArrayFromContents

    CollectZones = HasElements(outZones)
    Exit Function

ErrorHandler:
    CollectZones = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport.CollectZones"
End Function

' Safe bounds check (project pattern - mirrors ExportLengthInRegion.HasElements).
Private Function HasElements(ByRef arr() As Element) As Boolean
    On Error Resume Next
    HasElements = False
    If UBound(arr) <> -1 Then HasElements = True
    On Error GoTo 0
End Function

' ============================================================
'  PER-CABLE FIELD RESOLUTION
' ============================================================

' ResolveCableRepere
' Returns "<start label> - <end label>" using the Line's own geometric StartPoint/EndPoint order.
' A missing label on either end leaves that segment blank and marks bIncomplete.
Private Function ResolveCableRepere(ByVal oEl As Element, ByVal sRepereProp As String, _
                                    ByVal dRadius As Double, ByRef bIncomplete As Boolean) As String
    On Error GoTo ErrorHandler

    ResolveCableRepere = ""
    If Len(sRepereProp) = 0 Then Exit Function

    Dim ptStart As Point3d
    Dim ptEnd   As Point3d
    If Not GetCableEndpoints(oEl, ptStart, ptEnd) Then Exit Function   ' closed type: no two ends

    Dim sStart As String
    Dim sEnd   As String
    sStart = ResolveEndCellLabel(ptStart, sRepereProp, dRadius, bIncomplete)
    sEnd = ResolveEndCellLabel(ptEnd, sRepereProp, dRadius, bIncomplete)

    ResolveCableRepere = sStart & " - " & sEnd
    Exit Function

ErrorHandler:
    ResolveCableRepere = ""
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport.ResolveCableRepere"
End Function

' GetCableEndpoints
' Returns True + the two endpoints for an OPEN candidate type (Line/Arc/LineString/ComplexString).
' False for Shape/ComplexShape (closed - no two distinct ends to mark with a Repere) or any other
' type. LineString goes through the VertexList interface (Set oVL = oEl; GetVertices; first/last
' vertex) - the exact idiom Length.GetPartialLengthInsideZones already uses for this same type.
' ComplexString's StartPoint/EndPoint come from the ChainableElement interface it implements
' (confirmed: mvba-docs\02-objects\ChainableElement_Object.md documents the pair together; only
' StartPoint was previously exercised in this codebase, via Length.bas's own ComplexString branch).
Private Function GetCableEndpoints(ByVal oEl As Element, ByRef ptStart As Point3d, ByRef ptEnd As Point3d) As Boolean
    On Error GoTo ErrorHandler

    GetCableEndpoints = False

    Dim oVL     As VertexList
    Dim verts() As Point3d

    Select Case oEl.Type
        Case msdElementTypeLine
            ptStart = oEl.AsLineElement.StartPoint
            ptEnd = oEl.AsLineElement.EndPoint
        Case msdElementTypeArc
            ptStart = oEl.AsArcElement.StartPoint
            ptEnd = oEl.AsArcElement.EndPoint
        Case msdElementTypeLineString
            Set oVL = oEl
            verts = oVL.GetVertices
            ptStart = verts(LBound(verts))
            ptEnd = verts(UBound(verts))
        Case msdElementTypeComplexString
            ptStart = oEl.AsComplexStringElement.StartPoint
            ptEnd = oEl.AsComplexStringElement.EndPoint
        Case Else
            Exit Function   ' Shape/ComplexShape: closed, no two distinct ends
    End Select

    GetCableEndpoints = True
    Exit Function

ErrorHandler:
    GetCableEndpoints = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport.GetCableEndpoints"
End Function

' ResolveEndCellLabel
' Finds the nearest CellHeader within dRadius of pt that carries sRepereProp (via
' GetElements.FindNearestElement's own RequirePropertyName filter) and returns its value.
' No qualifying cell within radius: "" + bIncomplete - an expected drawing gap, never logged.
Private Function ResolveEndCellLabel(ByRef pt As Point3d, ByVal sRepereProp As String, _
                                     ByVal dRadius As Double, ByRef bIncomplete As Boolean) As String
    On Error GoTo ErrorHandler

    ResolveEndCellLabel = ""

    Dim oCell As Element
    Set oCell = GetElements.FindNearestElement(pt, dRadius, ElTypes:=Array(msdElementTypeCellHeader), RequirePropertyName:=sRepereProp)
    If oCell Is Nothing Then
        bIncomplete = True
        Exit Function
    End If

    Dim vVal As Variant
    vVal = CustomPropertyHandler.GetPropertyValueFromElement(oCell, sRepereProp, sRepereProp)
    If Not IsNull(vVal) Then ResolveEndCellLabel = Trim(CStr(vVal))
    Exit Function

ErrorHandler:
    ResolveEndCellLabel = ""
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport.ResolveEndCellLabel"
End Function

' ResolveCableText
' Scans oEl's WHOLE graphic group (Link.GetLink, no type filter) and reads Nature/Longueur off
' whichever member(s) carry them, verbatim (no parsing/reformatting - the values are written to
' Excel as-is). Each property is resolved INDEPENDENTLY (FindGroupPropertyValue), so they need not
' share the same element. No group, or either value absent from every member: Null + bIncomplete -
' an expected drawing gap, never logged.
Private Sub ResolveCableText(ByVal oEl As Element, ByVal sNatureProp As String, ByVal sLongueurProp As String, _
                             ByRef vNature As Variant, ByRef vLongueur As Variant, ByRef bIncomplete As Boolean)
    On Error GoTo ErrorHandler

    vNature = Null
    vLongueur = Null

    Dim linked() As Element
    linked = Link.GetLink(oEl)
    If Not HasElements(linked) Then
        bIncomplete = True
        Exit Sub
    End If

    If Len(sNatureProp) > 0 Then vNature = FindGroupPropertyValue(linked, sNatureProp)
    If Len(sLongueurProp) > 0 Then vLongueur = FindGroupPropertyValue(linked, sLongueurProp)
    If IsNull(vNature) Or IsNull(vLongueur) Then bIncomplete = True
    Exit Sub

ErrorHandler:
    bIncomplete = True
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport.ResolveCableText"
End Sub

' FindGroupPropertyValue
' Returns the first non-Null value of sPropName found scanning linked() in order. Nature/Longueur
' may live on any graphic-group member (a Cell, a Text, ...), not necessarily a specific element
' type, so this never filters by type - mirrors PropertyCalculation's own GROUP-source scans.
Private Function FindGroupPropertyValue(ByRef linked() As Element, ByVal sPropName As String) As Variant
    On Error GoTo ErrorHandler

    FindGroupPropertyValue = Null
    Dim i    As Long
    Dim vVal As Variant
    For i = LBound(linked) To UBound(linked)
        vVal = CustomPropertyHandler.GetPropertyValueFromElement(linked(i), sPropName, sPropName)
        If Not IsNull(vVal) Then
            FindGroupPropertyValue = vVal
            Exit Function
        End If
    Next i
    Exit Function

ErrorHandler:
    FindGroupPropertyValue = Null
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport.FindGroupPropertyValue"
End Function

' ============================================================
'  SOIL-TYPE PIVOT ACCUMULATION
' ============================================================

' RecordZoneLengths
' Measures this cable against every LABELLED zone (Length.GetPartialLengthInsideZones, 1-element array -
' the exact per-zone idiom ExportLengthInRegion.AggregateByZoneAndProperty already uses) and records each
' > 0 result raw, in oZoneCable, keyed zoneIdx & KEY_SEP & cableKey (one entry per pair by construction).
' NOTHING is aggregated here on purpose: a column key depends on how many cables share the zone, which is
' only known once every cable has been measured - BuildPivotAndSharedTrenches does the aggregation.
' Labels come pre-resolved from BuildZoneLabels; a blank label skips the zone entirely, geometry call
' included (a data-quality gap in the drawing, not worth a status flicker per occurrence).
Private Sub RecordZoneLengths(ByVal oEl As Element, ByVal sCableKey As String, ByRef zones() As Element, _
                              ByRef zoneLabels() As String, ByRef oZoneCable As Object)
    On Error GoTo ErrorHandler

    If Not HasElements(zones) Then Exit Sub

    Dim oneZone(0 To 0) As Element
    Dim i               As Long
    Dim z               As Long
    Dim dLen            As Double

    For i = LBound(zones) To UBound(zones)
        z = i - LBound(zones)
        If Len(zoneLabels(z)) > 0 Then
            Set oneZone(0) = zones(i)
            dLen = Length.GetPartialLengthInsideZones(oEl, oneZone)
            If dLen > 0 Then oZoneCable.Add CStr(z) & KEY_SEP & sCableKey, dLen
        End If
    Next i
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport.RecordZoneLengths"
End Sub

' BuildZoneLabels
' Fills outLabels (0-based, indexed by zone scan position) with each zone's soil-type value, read ONCE
' (Trim(CStr)); "" when the property is absent/Null/blank. Unlike ExportLengthInRegion.BuildZoneLabels
' there is deliberately NO "Zone <n>" fallback: here a blank label means "this zone contributes to
' nothing". Returns the zone count (0 on an empty array or a fault).
Private Function BuildZoneLabels(ByRef zones() As Element, ByVal sZoneProp As String, ByRef outLabels() As String) As Long
    On Error GoTo ErrorHandler

    BuildZoneLabels = 0
    If Not HasElements(zones) Then Exit Function

    Dim i    As Long
    Dim z    As Long
    Dim vVal As Variant
    ReDim outLabels(0 To UBound(zones) - LBound(zones))
    For i = LBound(zones) To UBound(zones)
        z = i - LBound(zones)
        outLabels(z) = ""
        vVal = CustomPropertyHandler.GetPropertyValueFromElement(zones(i), sZoneProp, sZoneProp)
        If Not IsNull(vVal) Then outLabels(z) = Trim(CStr(vVal))
    Next i

    BuildZoneLabels = UBound(zones) - LBound(zones) + 1
    Exit Function

ErrorHandler:
    BuildZoneLabels = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport.BuildZoneLabels"
End Function

' BuildPivotAndSharedTrenches
' Post-pass over the per-(zone, cable) lengths recorded by RecordZoneLengths. For every labelled zone,
' the cables inside it (len > 0) are collected in row order, and the zone gets its COLUMN KEY:
'   1 cable  -> "<label>"        (an ordinary stretch: this cable digs it alone; all solo zones of that
'                                 soil type share the one column)
'   n cables -> "<label> (n)"    (a SHARED TRENCH: one trench, n cables in it - e.g. "CH2C (2)"), keyed
'                                 on the SHARING GROUP: zones with the same members merge (two stretches
'                                 of the same shared trench add up), zones with the same label and count
'                                 but DIFFERENT members do not - "TA1 (2)" for A+B and "TA1 (2) #2" for
'                                 A+C are two distinct trenches and stay two distinct columns.
' So a shared trench reads as its own column next to the ordinary one, instead of silently inflating it.
' Several shared columns can coexist ("CH2C", "CH2C (2)", "CH2C (3)", "TRIC (2)"), and one cable can
' appear in several of them - a cable can share its start with one cable and its end with another.
' Per column key: oPivot(cable) += that cable's own length (rows stay cable-side truth), oCableTotal +=
' every cable's length (what the rows add up to) and oTrench += the MAX length in the zone (the trench is
' dug ONCE, along its longest occupant - one zone = one trench segment; see the module header for the
' one-common-zone assumption). A shared zone also cross-marks each cable with the OTHER cables' Reperes
' in oShared (deduplicated across zones, ", "-joined). Returns the number of shared zones.
Private Function BuildPivotAndSharedTrenches(ByRef oZoneCable As Object, ByRef zoneLabels() As String, ByVal nZones As Long, _
                                             ByRef rowOrder() As String, ByRef oRowData As Object, _
                                             ByRef oColumns As Object, ByRef oPivot As Object, _
                                             ByRef oShared As Object, ByRef oCableTotal As Object, ByRef oTrench As Object) As Long
    On Error GoTo ErrorHandler

    BuildPivotAndSharedTrenches = 0
    If nZones = 0 Then Exit Function

    Dim z         As Long
    Dim r         As Long
    Dim k         As Long
    Dim sLabel    As String
    Dim sColKey   As String
    Dim sBase     As String
    Dim sGroupKey As String
    Dim sKey      As String
    Dim sPivotKey As String
    Dim dLen      As Double
    Dim dMax      As Double
    Dim dSum      As Double
    Dim inZone()  As Long      ' rowOrder indexes of the cables inside the current zone (ascending)
    Dim inLen()   As Double    ' their lengths, same order
    Dim nIn       As Long
    Dim nShared   As Long
    Dim oGroupCol As Object    ' group key (label + member set) -> column key, so one sharing group = one column
    Dim oBaseCount As Object   ' "<label> (n)" -> how many DISTINCT groups already carry it (suffix source)

    Set oGroupCol = CreateObject("Scripting.Dictionary")
    Set oBaseCount = CreateObject("Scripting.Dictionary")
    nShared = 0
    For z = 0 To nZones - 1
        sLabel = zoneLabels(z)
        If Len(sLabel) > 0 Then
            nIn = 0
            dMax = 0
            dSum = 0
            ReDim inZone(0 To UBound(rowOrder) - LBound(rowOrder))
            ReDim inLen(0 To UBound(rowOrder) - LBound(rowOrder))
            For r = LBound(rowOrder) To UBound(rowOrder)
                sKey = CStr(z) & KEY_SEP & rowOrder(r)
                If oZoneCable.Exists(sKey) Then
                    dLen = oZoneCable(sKey)
                    inZone(nIn) = r
                    inLen(nIn) = dLen
                    nIn = nIn + 1
                    dSum = dSum + dLen
                    If dLen > dMax Then dMax = dLen
                End If
            Next r

            If nIn > 0 Then
                If nIn = 1 Then
                    ' Ordinary stretch: every solo zone of this soil type shares the plain column.
                    sColKey = sLabel
                Else
                    ' Shared trench: the column identifies the SHARING GROUP, not just the count - two
                    ' TA1 zones each shared by 2 cables but by DIFFERENT pairs (A+B here, A+C there) are
                    ' two different trenches and must not merge. Group key = label + the member set (row
                    ' indexes, ascending, so it is canonical); zones with the SAME members do merge, which
                    ' is wanted (two stretches of the same shared trench add up). Same label AND same count
                    ' but a different group -> the column carries a " #k" suffix, numbered in zone scan
                    ' order; which cables belong to it is readable from the column's own filled rows.
                    sGroupKey = sLabel & KEY_SEP
                    For k = 0 To nIn - 1
                        sGroupKey = sGroupKey & CStr(inZone(k)) & "|"
                    Next k

                    If oGroupCol.Exists(sGroupKey) Then
                        sColKey = oGroupCol(sGroupKey)
                    Else
                        sBase = sLabel & " (" & nIn & ")"
                        If oBaseCount.Exists(sBase) Then
                            oBaseCount(sBase) = oBaseCount(sBase) + 1
                        Else
                            oBaseCount.Add sBase, 1
                        End If
                        If oBaseCount(sBase) = 1 Then
                            sColKey = sBase
                        Else
                            sColKey = sBase & " #" & oBaseCount(sBase)
                        End If
                        oGroupCol.Add sGroupKey, sColKey
                    End If
                End If
                If Not oColumns.Exists(sColKey) Then oColumns.Add sColKey, True

                For k = 0 To nIn - 1
                    sPivotKey = rowOrder(inZone(k)) & KEY_SEP & sColKey
                    If oPivot.Exists(sPivotKey) Then
                        oPivot(sPivotKey) = oPivot(sPivotKey) + inLen(k)
                    Else
                        oPivot.Add sPivotKey, inLen(k)
                    End If
                Next k

                If oCableTotal.Exists(sColKey) Then
                    oCableTotal(sColKey) = oCableTotal(sColKey) + dSum
                Else
                    oCableTotal.Add sColKey, dSum
                End If
                If oTrench.Exists(sColKey) Then
                    oTrench(sColKey) = oTrench(sColKey) + dMax
                Else
                    oTrench.Add sColKey, dMax
                End If
            End If

            If nIn >= 2 Then
                nShared = nShared + 1
                For r = 0 To nIn - 1
                    For k = 0 To nIn - 1
                        If k <> r Then
                            AppendSharedMark oShared, rowOrder(inZone(r)), CableIdentifier(rowOrder(inZone(k)), inZone(k), oRowData)
                        End If
                    Next k
                Next r
            End If
        End If
    Next z

    BuildPivotAndSharedTrenches = nShared
    Exit Function

ErrorHandler:
    BuildPivotAndSharedTrenches = 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport.BuildPivotAndSharedTrenches"
End Function

' CableIdentifier
' What a cable is called in another cable's "Shared trench" cell: its Repere when non-blank, else
' "#<Excel row>" (rowOrder index + 2, row 1 being the header) - a stable, human-readable pointer even
' for a cable with no marker cells. A Repere with both ends blank renders as "-" and counts as blank.
Private Function CableIdentifier(ByVal sCableKey As String, ByVal nRowIdx As Long, ByRef oRowData As Object) As String
    On Error GoTo ErrorHandler

    Dim vRow As Variant
    vRow = oRowData(sCableKey)
    CableIdentifier = Trim(CStr(vRow(0)))
    If Len(CableIdentifier) = 0 Then CableIdentifier = "#" & (nRowIdx + 2)
    If CableIdentifier = "-" Then CableIdentifier = "#" & (nRowIdx + 2)
    Exit Function

ErrorHandler:
    CableIdentifier = "#" & (nRowIdx + 2)
End Function

' AppendSharedMark
' Appends sMark to oShared(sCableKey) as a ", "-joined list, unless it is already listed (the same
' pair of cables can share several zones - listed once).
Private Sub AppendSharedMark(ByRef oShared As Object, ByVal sCableKey As String, ByVal sMark As String)
    On Error GoTo ErrorHandler

    Dim sCur As String
    sCur = ""
    If oShared.Exists(sCableKey) Then sCur = oShared(sCableKey)
    If InStr(1, ", " & sCur & ", ", ", " & sMark & ", ", vbTextCompare) > 0 Then Exit Sub

    If Len(sCur) > 0 Then
        sCur = sCur & ", " & sMark
    Else
        sCur = sMark
    End If
    If oShared.Exists(sCableKey) Then
        oShared(sCableKey) = sCur
    Else
        oShared.Add sCableKey, sCur
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport.AppendSharedMark"
End Sub

' ============================================================
'  CONFIGURED PROPERTY-NAME RESOLUTION (mirrors ExportLengthInRegion.ResolveZoneLabelProperty)
' ============================================================

' ResolveConfiguredProperty
' empty                          -> "" (silent; that column stays blank / that zone gets no label)
' no ItemType enumerable at all  -> "" + informational log, NO status (DGNLib absent/empty)
' non-empty, an ARES ItemType    -> the name
' non-empty, NOT an ItemType     -> log + one-shot CableReportPropertyInvalid status + ""
Private Function ResolveConfiguredProperty(ByVal sConfiguredName As String, ByVal sFieldLabel As String) As String
    On Error GoTo ErrorHandler

    ResolveConfiguredProperty = ""
    Dim sName As String
    sName = Trim(sConfiguredName)
    If Len(sName) = 0 Then Exit Function

    Dim names() As String
    names = CustomPropertyHandler.GetCustomPropertyNames()

    If IsEmptyNameList(names) Then
        ErrorHandler.HandleError "Property '" & sName & "' (" & sFieldLabel & ") left unresolved: no ARES ItemType could be enumerated (DGNLib absent or empty)", 0, "", "CableReport.ResolveConfiguredProperty"
        Exit Function
    End If

    If NameInList(sName, names) Then
        ResolveConfiguredProperty = sName
        Exit Function
    End If

    ErrorHandler.HandleError "Property configured but not an ItemType of the ARES DGNLib: '" & sName & "' (" & sFieldLabel & ")", 0, "", "CableReport.ResolveConfiguredProperty"
    ShowStatus GetTranslation("CableReportPropertyInvalid", sFieldLabel, sName)
    Exit Function

ErrorHandler:
    ResolveConfiguredProperty = ""
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport.ResolveConfiguredProperty"
End Function

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
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport.NameInList"
End Function

Private Function IsEmptyNameList(ByRef names() As String) As Boolean
    On Error GoTo ErrorHandler

    IsEmptyNameList = True
    Dim i As Long
    For i = LBound(names) To UBound(names)
        If Len(Trim(names(i))) > 0 Then
            IsEmptyNameList = False
            Exit Function
        End If
    Next i
    Exit Function

ErrorHandler:
    IsEmptyNameList = True
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport.IsEmptyNameList"
End Function

' ============================================================
'  EXCEL EXPORT
' ============================================================

Private Function BuildDefaultFilename() As String
    BuildDefaultFilename = "ARES_CableReport_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"
End Function

' WriteToExcel
' Late-binds Excel, writes fixed headers (Repere/Nature/Longueur/Shared with) + one alphabetically-sorted
' pivot column per distinct column key - "<label>" for ordinary stretches, "<label> (n)" for a shared
' trench (see BuildPivotAndSharedTrenches); sorting keeps "CH2C" and "CH2C (2)" adjacent. One row per
' cable in scan order (rowOrder - no natural sort key for a Repere string). A pivot cell with no
' contribution stays BLANK, never 0. When there are pivot columns, a two-row footer follows a blank row:
' "Cable totals" (plain per-column sums, what the rows add up to) and "Trench to dig (deduplicated)"
' (each zone counted once, at its longest occupant).
'
' COM lifecycle: identical contract to ExportLengthInRegion.WriteToExcel (bExcelStartedByUs,
' always-attempted close on error, Quit gated on having started the session ourselves).
Private Sub WriteToExcel(ByRef oRowData As Object, ByRef oColumns As Object, ByRef oPivot As Object, _
                         ByRef oShared As Object, ByRef oCableTotal As Object, ByRef oTrench As Object, _
                         ByRef rowOrder() As String, ByVal Filepath As String, ByVal bVisible As Boolean)

    Dim xlApp             As Object
    Dim xlBook            As Object
    Dim xlSheet           As Object
    Dim bExcelStartedByUs As Boolean
    Dim nRound            As Byte
    Dim colKeys()         As String
    Dim i                 As Long
    Dim c                 As Long
    Dim r                 As Long
    Dim rFoot             As Long
    Dim sKey              As String
    Dim sPivotKey         As String
    Dim vRow              As Variant
    Dim dNum              As Double

    On Error GoTo ErrorHandler

    Dim sRound As String
    sRound = ARESConfig.ARES_CABLEREPORT_ROUND.Value
    If Len(sRound) = 0 Then sRound = ARESConfig.ARES_CABLEREPORT_ROUND.DefaultValue
    nRound = CByte(sRound)
    If nRound = ARES_RND_ERROR_VALUE Then nRound = CByte(ARESConfig.ARES_CABLEREPORT_ROUND.DefaultValue)

    ' (1) Reuse existing Excel session if the user already has one; never quit it on cleanup.
    bExcelStartedByUs = True
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    On Error GoTo ErrorHandler
    If Not xlApp Is Nothing Then
        bExcelStartedByUs = False
    Else
        Set xlApp = CreateObject("Excel.Application")
    End If

    ' (2) New workbook, name the sheet.
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    xlSheet.Name = GetTranslation("CableReportSheetName")

    ' (3) Fixed headers (4) + one sorted column per distinct soil-type label from column 5.
    xlSheet.Cells(1, 1).Value = GetTranslation("CableReportHeaderRepere")
    xlSheet.Cells(1, 2).Value = GetTranslation("CableReportHeaderNature")
    xlSheet.Cells(1, 3).Value = GetTranslation("CableReportHeaderLongueur")
    xlSheet.Cells(1, 4).Value = GetTranslation("CableReportHeaderSharedWith")

    If oColumns.Count > 0 Then colKeys = SortedKeysCI(oColumns)
    For c = 0 To oColumns.Count - 1
        xlSheet.Cells(1, 5 + c).Value = colKeys(c)
    Next c

    ' (4) Data rows, one per cable, in scan order.
    For i = LBound(rowOrder) To UBound(rowOrder)
        sKey = rowOrder(i)
        r = i - LBound(rowOrder) + 2
        vRow = oRowData(sKey)
        xlSheet.Cells(r, 1).Value = vRow(0)
        If Not IsNull(vRow(1)) Then xlSheet.Cells(r, 2).Value = vRow(1)
        If Not IsNull(vRow(2)) Then
            ' Longueur: write a real Double when the property's value is purely numeric, so the column
            ' is not half text / half number (see TryAsNumber and the module header). Anything else -
            ' a unit suffix, a free-text value - stays verbatim.
            If TryAsNumber(vRow(2), dNum) Then
                xlSheet.Cells(r, 3).Value = dNum
            Else
                xlSheet.Cells(r, 3).Value = vRow(2)
            End If
        End If
        If oShared.Exists(sKey) Then xlSheet.Cells(r, 4).Value = oShared(sKey)
        For c = 0 To oColumns.Count - 1
            sPivotKey = sKey & KEY_SEP & colKeys(c)
            If oPivot.Exists(sPivotKey) Then xlSheet.Cells(r, 5 + c).Value = Round(oPivot(sPivotKey), nRound)
        Next c
    Next i

    ' (4b) Footer - only when there are soil-type columns: one blank row, then the plain per-label sums
    '      (what the cable rows add up to) and the deduplicated trench length (each zone once, at its
    '      longest occupant). The gap between the two rows IS the shared-trench overlap.
    If oColumns.Count > 0 Then
        rFoot = (UBound(rowOrder) - LBound(rowOrder) + 2) + 2   ' last data row + blank row + 1
        xlSheet.Cells(rFoot, 1).Value = GetTranslation("CableReportFooterCableTotal")
        xlSheet.Cells(rFoot + 1, 1).Value = GetTranslation("CableReportFooterTrenchTotal")
        For c = 0 To oColumns.Count - 1
            If oCableTotal.Exists(colKeys(c)) Then xlSheet.Cells(rFoot, 5 + c).Value = Round(oCableTotal(colKeys(c)), nRound)
            If oTrench.Exists(colKeys(c)) Then xlSheet.Cells(rFoot + 1, 5 + c).Value = Round(oTrench(colKeys(c)), nRound)
        Next c
    End If

    ' (5) Save when a path is provided.
    If Len(Filepath) > 0 Then
        xlApp.DisplayAlerts = False
        xlBook.SaveAs Filepath, XL_OPENXML_FORMAT
        xlApp.DisplayAlerts = True
    End If

    ' (6) Surface the workbook or close the headless session.
    If bVisible Then
        xlApp.Visible = True
    Else
        On Error Resume Next
        xlBook.Close False
        If bExcelStartedByUs Then xlApp.Quit
        On Error GoTo ErrorHandler
    End If

    On Error Resume Next
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport.WriteToExcel"
    On Error Resume Next
    If Not xlApp Is Nothing Then xlApp.DisplayAlerts = True
    If Not xlBook Is Nothing Then xlBook.Close False
    If bExcelStartedByUs Then
        If Not xlApp Is Nothing Then xlApp.Quit
    End If
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
End Sub

' TryAsNumber
' True + dOut when v is a purely numeric value: an already-numeric Variant, or a String that parses
' cleanly - first with the station's OWN locale (IsNumeric/CDbl are locale-aware, so "493,6" parses on a
' French host), then with the decimal separator swapped, so a value authored on a station using the other
' convention still lands as a number. False for anything else (a unit suffix like "493,6m", an identifier,
' Null/Empty/an array) - the caller then writes the value verbatim.
' WHY: assigning a STRING to Excel's Range.Value goes through the object model's INVARIANT (en-US)
' parsing, not the user's locale, so "493,6" is stored as TEXT; assigning a Double never is. That is what
' made the Longueur column text while every ARES-computed cell was numeric.
Private Function TryAsNumber(ByVal v As Variant, ByRef dOut As Double) As Boolean
    On Error GoTo ErrorHandler

    TryAsNumber = False
    dOut = 0

    ' Type tests FIRST, each on its own line: VBA has no short-circuit, and CStr on an array raises 13.
    If IsArray(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    Select Case VarType(v)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbByte, vbCurrency, vbDecimal
            dOut = CDbl(v)
            TryAsNumber = True
            Exit Function
        Case vbString
            ' handled below
        Case Else
            Exit Function                  ' Boolean/Date/Object are never coerced into a length here
    End Select

    Dim s As String
    s = Trim(CStr(v))
    If Len(s) = 0 Then Exit Function

    If IsNumeric(s) Then
        dOut = CDbl(s)
        TryAsNumber = True
        Exit Function
    End If

    ' Separator swap - only when exactly ONE of the two marks is present, so a thousands-grouped value
    ' ("1.234,5") is never silently reinterpreted; it simply stays text.
    Dim sSwapped As String
    sSwapped = ""
    If InStr(s, ",") > 0 Then
        If InStr(s, ".") = 0 Then sSwapped = Replace(s, ",", ".")
    ElseIf InStr(s, ".") > 0 Then
        sSwapped = Replace(s, ".", ",")
    End If

    If Len(sSwapped) > 0 Then
        If IsNumeric(sSwapped) Then
            dOut = CDbl(sSwapped)
            TryAsNumber = True
        End If
    End If
    Exit Function

ErrorHandler:
    TryAsNumber = False
    dOut = 0
End Function

' SortedKeysCI
' Returns a 0-based String() of the dictionary keys sorted case-insensitive (mirrors
' ExportLengthInRegion.SortedKeysCI). Column counts are small - bubble sort is fine.
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
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport.SortedKeysCI"
End Function
