' CableReport.bas
' Description: One Excel row per cable (levels in ARES_CableReport_Cable_Level): its end-cell Repere
'              markers, the Nature/Longueur carried by its graphic group, and its trenching length
'              pivoted per soil-type zone value - a zone shared by several cables getting its own
'              "<value> (n)" column. Column model, shared-trench rules and the per-field resolution
'              doctrine: see MVBA/README.md (Cable Report) and the wiki page of the same name.
' ENTRY POINT  CableReport([CableLevel], [ZoneLevel], [Filepath], [ExcelVisible]) - ZoneLevel empty
'              falls back to ARES_Outline_Output_Level, Filepath empty opens a Save-As (cancel aborts).
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

    ' --- Zone labels resolved ONCE (not per cable x zone); "" = the zone contributes to nothing ---
    Dim zoneLabels() As String
    Dim nZones       As Long
    nZones = 0
    If bHasZones And Len(sZoneProp) > 0 Then nZones = BuildZoneLabels(zones, sZoneProp, zoneLabels)

    ' --- One row per cable + the raw per-(zone, cable) lengths; the pivot columns cannot be decided
    ' here (a shared-trench key needs the zone's cable count), so BuildPivotAndSharedTrenches does it ---
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
' True + the two endpoints for an OPEN candidate type (Line/Arc/LineString/ComplexString); False for a
' closed Shape/ComplexShape (no two distinct ends to mark) or any other type. LineString goes through
' the VertexList interface, ComplexString through ChainableElement's StartPoint/EndPoint.
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
' Reads Nature/Longueur verbatim off whichever member of oEl's graphic group carries them, each
' resolved INDEPENDENTLY (they need not share an element). Absent: Null + bIncomplete, never logged -
' an expected drawing gap.
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
' Measures this cable against every LABELLED zone (Length.GetPartialLengthInsideZones with a 1-element
' array, the per-zone idiom ExportLengthInRegion already uses) and records each > 0 result raw, keyed
' zoneIdx & KEY_SEP & cableKey. NOTHING is aggregated here on purpose: a column key depends on how many
' cables share the zone, which is only known once every cable has been measured. A blank label skips the
' zone, geometry call included.
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
' Post-pass over the per-(zone, cable) lengths recorded by RecordZoneLengths: derives the pivot columns,
' the pivot itself, the shared-trench cross-marks and the two footer series (see MVBA/README.md for the
' column model). Column key = "<label>" for a solo zone, "<label> (n)" for a shared one, keyed on the
' sharing GROUP so two same-count zones with different members stay two columns. Returns the number of
' shared zones. ASSUMPTION: a shared trench is ONE common zone - two overlapping zones are not detected.
' oTrench takes the MAX length per zone, not the sum: a zone is one trench segment, dug once.
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
                    ' Group key = label + member set (row indexes, ascending, hence canonical): same
                    ' members merge, different members get their own column - suffixed " #k" when the
                    ' base "<label> (n)" is already taken by another group, numbered in zone scan order.
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
' 4 fixed headers (Repere/Nature/Longueur/Shared with) + one alphabetically-sorted pivot column per
' column key (sorting keeps "CH2C" and "CH2C (2)" adjacent), one row per cable in scan order, then the
' two footer rows when there are pivot columns. A pivot cell with no contribution stays BLANK, never 0.
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
            ' Longueur as a real Double when it is purely numeric, else verbatim - see TryAsNumber.
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
' True + dOut when v is purely numeric: an already-numeric Variant, or a String that parses cleanly -
' with the station's OWN locale first (IsNumeric/CDbl are locale-aware), then with the decimal separator
' swapped. False for anything else (a unit suffix like "493,6m", an identifier, Null/Empty/an array).
' WHY it exists: a STRING assigned to Excel's Range.Value is parsed with the object model's INVARIANT
' locale, so "493,6" lands as TEXT; a Double never does.
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

' ============================================================
'  DIAGNOSTIC
' ============================================================

' Prints what every trench column is built from, so the two numbers that disagree in the sheet can be
' traced to their source. Read-only: it measures and prints, it changes nothing.
'
' The sheet's "Longueur" is the cable's own Longueur PROPERTY, while the trench columns are GEOMETRY
' measured zone by zone. They share no origin, so when they disagree only these four figures say which
' one is wrong:
'   geom total   - the cable's real length (Length.GetLength). Compare it to the property: a gap here
'                  means the property is stale or was calculated from something else.
'   sum of zones - what the columns add up to, one GetPartialLengthInsideZones call per zone, exactly
'                  as RecordZoneLengths does it.
'   union        - the same call with EVERY labelled zone at once. PointInAnyZone treats the array as
'                  a union, so a portion crossed by two OVERLAPPING zones counts once here and twice
'                  in the sum. sum > union is therefore proof that zones overlap.
'   per end      - what each cable end resolved to, with its coordinates, for a Repere that came out
'                  empty on one side.
Public Sub DiagCableLengths()
    On Error GoTo ErrorHandler

    If Not Application.HasActiveModelReference Then
        Debug.Print "DIAG: no active model reference."
        Exit Sub
    End If

    Dim sCableLevel As String
    Dim sZoneLevel  As String
    sCableLevel = ARESConfig.ARES_CABLEREPORT_CABLE_LEVEL.value
    sZoneLevel = ARESConfig.ARES_CABLEREPORT_ZONE_LEVEL.value
    If Len(sZoneLevel) = 0 Then sZoneLevel = ARESConfig.ARES_OUTLINE_OUTPUT_LEVEL.value

    Dim cableLevels()  As String
    Dim sIgnored       As String
    If ResolveCableLevels(sCableLevel, cableLevels, sIgnored) = 0 Then
        Debug.Print "DIAG: no usable cable level in '" & sCableLevel & "'."
        Exit Sub
    End If

    Dim sRepereProp   As String
    Dim sLongueurProp As String
    Dim sZoneProp     As String
    sRepereProp = ResolveConfiguredProperty(ARESConfig.ARES_CABLEREPORT_REPERE_PROPERTY.value, "Repere")
    sLongueurProp = ResolveConfiguredProperty(ARESConfig.ARES_CABLEREPORT_LONGUEUR_PROPERTY.value, "Longueur")
    sZoneProp = ResolveConfiguredProperty(ARESConfig.ARES_CABLEREPORT_ZONE_PROPERTY.value, "Coupe_Type")

    Dim cables() As Element
    If Not CollectCables(cableLevels, cables) Then
        Debug.Print "DIAG: no cable found on " & sCableLevel
        Exit Sub
    End If

    Dim zones()      As Element
    Dim zoneLabels() As String
    Dim nZones       As Long
    nZones = 0
    If CollectZones(sZoneLevel, zones) And Len(sZoneProp) > 0 Then
        nZones = BuildZoneLabels(zones, sZoneProp, zoneLabels)
    End If

    ' Every labelled zone in one array - the union reference the per-zone sum is compared against.
    Dim labelled() As Element
    Dim nLab       As Long
    nLab = 0
    If nZones > 0 Then
        ReDim labelled(0 To nZones - 1)
        Dim k As Long
        For k = LBound(zones) To UBound(zones)
            If Len(zoneLabels(k - LBound(zones))) > 0 Then
                Set labelled(nLab) = zones(k)
                nLab = nLab + 1
            End If
        Next k
        If nLab > 0 Then ReDim Preserve labelled(0 To nLab - 1)
    End If

    Dim dRadius As Double
    dRadius = Val(ARESConfig.ARES_CABLEREPORT_SEARCH_RADIUS.value)
    If dRadius <= 0 Then dRadius = Val(ARESConfig.ARES_CABLEREPORT_SEARCH_RADIUS.DefaultValue)

    Debug.Print String(78, "=")
    Debug.Print "CABLE REPORT DIAG - " & (UBound(cables) - LBound(cables) + 1) & " cable(s), " & _
                nLab & " labelled zone(s), radius " & dRadius
    Debug.Print String(78, "=")

    Dim i        As Long
    Dim oEl      As Element
    Dim oneZone(0 To 0) As Element
    Dim z        As Long
    Dim dLen     As Double
    Dim dSum     As Double
    Dim dUnion   As Double
    Dim dGeom    As Double
    Dim ptS      As Point3d
    Dim ptE      As Point3d
    Dim bInc     As Boolean
    Dim vProp    As Variant
    Dim sProp    As String
    Dim linked() As Element

    For i = LBound(cables) To UBound(cables)
        Set oEl = cables(i)
        Debug.Print "cable #" & (i - LBound(cables) + 1) & "  id=" & DLongToString(oEl.ID) & "  type=" & oEl.Type

        ' --- ends / Repere ---
        If GetCableEndpoints(oEl, ptS, ptE) Then
            bInc = False
            Debug.Print "   end A " & Format(ptS.X, "0.000") & ";" & Format(ptS.Y, "0.000") & _
                        "  -> [" & ResolveEndCellLabel(ptS, sRepereProp, dRadius, bInc) & "]"
            Debug.Print "   end B " & Format(ptE.X, "0.000") & ";" & Format(ptE.Y, "0.000") & _
                        "  -> [" & ResolveEndCellLabel(ptE, sRepereProp, dRadius, bInc) & "]"
        Else
            Debug.Print "   ends   : closed type, no two ends"
        End If

        ' --- property vs geometry ---
        ' Read it the way the export does - across the graphic GROUP, not on the cable itself. The
        ' tagging rules attach Longueur to a linked element, so asking the cable directly answers
        ' "<none>" on a drawing where the sheet shows the value perfectly well.
        vProp = Null
        linked = Link.GetLink(oEl)
        If HasElements(linked) Then vProp = FindGroupPropertyValue(linked, sLongueurProp)
        dGeom = Length.GetLength(oEl)
        sProp = "<none>"
        If Not IsArray(vProp) Then
            If Not IsNull(vProp) Then
                If Not IsEmpty(vProp) Then sProp = CStr(vProp)
            End If
        End If
        Debug.Print "   Longueur property : [" & sProp & "]"
        Debug.Print "   geom total        : " & Format(dGeom, "0.00")

        ' --- per zone, exactly as RecordZoneLengths measures ---
        dSum = 0
        If nZones > 0 Then
            For z = LBound(zones) To UBound(zones)
                If Len(zoneLabels(z - LBound(zones))) > 0 Then
                    Set oneZone(0) = zones(z)
                    dLen = Length.GetPartialLengthInsideZones(oEl, oneZone)
                    If dLen > 0 Then
                        dSum = dSum + dLen
                        Debug.Print "     zone " & (z - LBound(zones)) & " [" & zoneLabels(z - LBound(zones)) & "] : " & Format(dLen, "0.00")
                    End If
                End If
            Next z
        End If

        dUnion = 0
        If nLab > 0 Then dUnion = Length.GetPartialLengthInsideZones(oEl, labelled)

        Debug.Print "   sum of zones      : " & Format(dSum, "0.00")
        Debug.Print "   union of zones    : " & Format(dUnion, "0.00")
        If dSum - dUnion > 0.01 Then
            Debug.Print "   -> OVERLAP: the sum counts " & Format(dSum - dUnion, "0.00") & " m twice or more."
        End If
        If dUnion - dGeom > 0.01 Then
            Debug.Print "   -> union EXCEEDS the cable itself by " & Format(dUnion - dGeom, "0.00") & " m."
        End If
        Debug.Print
    Next i

    Debug.Print String(78, "=")
    Exit Sub

ErrorHandler:
    Debug.Print "DIAG ERROR " & Err.Number & " - " & Err.Description
End Sub

' Localises WHERE the per-zone sum and the union disagree, by growing the zone set one at a time.
' For each labelled zone in turn: its own measured length, and how much it actually ADDS to the union
' of everything before it. Equal means the zone contributes fresh cable. Less means that portion was
' already covered - which is either an overlap with an earlier zone, or a defect in the multi-zone
' path. Concentrated tiny deltas point at abutting zone joints; a few large ones point at something
' structural. Read-only, and O(n) calls on a growing array, so it takes a moment on many zones.
Public Sub DiagZoneOverlap()
    On Error GoTo ErrorHandler

    If Not Application.HasActiveModelReference Then
        Debug.Print "DIAG: no active model reference."
        Exit Sub
    End If

    Dim sCableLevel As String
    Dim sZoneLevel  As String
    sCableLevel = ARESConfig.ARES_CABLEREPORT_CABLE_LEVEL.value
    sZoneLevel = ARESConfig.ARES_CABLEREPORT_ZONE_LEVEL.value
    If Len(sZoneLevel) = 0 Then sZoneLevel = ARESConfig.ARES_OUTLINE_OUTPUT_LEVEL.value

    Dim cableLevels() As String
    Dim sIgnored      As String
    If ResolveCableLevels(sCableLevel, cableLevels, sIgnored) = 0 Then Exit Sub

    Dim sZoneProp As String
    sZoneProp = ResolveConfiguredProperty(ARESConfig.ARES_CABLEREPORT_ZONE_PROPERTY.value, "Coupe_Type")

    Dim cables() As Element
    If Not CollectCables(cableLevels, cables) Then Exit Sub

    Dim zones()      As Element
    Dim zoneLabels() As String
    Dim nZones       As Long
    nZones = 0
    If CollectZones(sZoneLevel, zones) And Len(sZoneProp) > 0 Then
        nZones = BuildZoneLabels(zones, sZoneProp, zoneLabels)
    End If
    If nZones = 0 Then
        Debug.Print "DIAG: no labelled zone."
        Exit Sub
    End If

    Dim oEl As Element
    Set oEl = cables(LBound(cables))

    Dim grow()   As Element
    Dim one(0 To 0) As Element
    Dim k        As Long
    Dim nUsed    As Long
    Dim dOwn     As Double
    Dim dUnion   As Double
    Dim dPrev    As Double
    Dim dInc     As Double
    Dim dSumOwn  As Double
    Dim dLost    As Double

    ReDim grow(0 To nZones - 1)
    nUsed = 0
    dPrev = 0
    dSumOwn = 0
    dLost = 0

    Debug.Print String(78, "=")
    Debug.Print "ZONE OVERLAP DIAG - cable id=" & DLongToString(oEl.ID) & ", geom " & Format(Length.GetLength(oEl), "0.00")
    Debug.Print "  zone : own   increment  lost"
    Debug.Print String(78, "=")

    For k = LBound(zones) To UBound(zones)
        If Len(zoneLabels(k - LBound(zones))) > 0 Then
            Set one(0) = zones(k)
            dOwn = Length.GetPartialLengthInsideZones(oEl, one)

            Set grow(nUsed) = zones(k)
            nUsed = nUsed + 1

            Dim probe() As Element
            ReDim probe(0 To nUsed - 1)
            Dim j As Long
            For j = 0 To nUsed - 1
                Set probe(j) = grow(j)
            Next j
            dUnion = Length.GetPartialLengthInsideZones(oEl, probe)

            dInc = dUnion - dPrev
            dPrev = dUnion
            dSumOwn = dSumOwn + dOwn
            If dOwn - dInc > 0.01 Then dLost = dLost + (dOwn - dInc)

            If dOwn > 0 Then
                Debug.Print "  " & Format(k - LBound(zones), "00") & " [" & zoneLabels(k - LBound(zones)) & "] : " & _
                            Format(dOwn, "0.00") & "   " & Format(dInc, "0.00") & _
                            IIf(dOwn - dInc > 0.01, "   LOST " & Format(dOwn - dInc, "0.00"), "")
            End If
        End If
    Next k

    Debug.Print String(78, "-")
    Debug.Print "sum of own lengths : " & Format(dSumOwn, "0.00")
    Debug.Print "final union        : " & Format(dPrev, "0.00")
    Debug.Print "accounted as lost  : " & Format(dLost, "0.00")
    Debug.Print String(78, "=")
    Exit Sub

ErrorHandler:
    Debug.Print "DIAG ERROR " & Err.Number & " - " & Err.Description
End Sub
