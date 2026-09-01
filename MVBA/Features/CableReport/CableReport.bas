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
' SOIL-TYPE BREAKDOWN   for each zone (Shape/ComplexShape on ARES_CableReport_Zone_Level) whose
'                        partial length inside it (Length.GetPartialLengthInsideZones) is > 0, the
'                        length is summed into a column keyed by that zone's configured property
'                        value (default "Coupe_Type"). Two zones sharing the same value for the
'                        same cable SUM into one column. Column order is alphabetical (case-
'                        insensitive), not scan order - deterministic run to run.
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

    ' --- Accumulate one row per cable + pivot columns ---
    Dim oRowData As Object   ' Scripting.Dictionary: cableKey -> Array(sRepere, vNature, vLongueur)
    Dim oColumns As Object   ' Scripting.Dictionary: zone label -> True (presence only; sorted at write time)
    Dim oPivot   As Object   ' Scripting.Dictionary: cableKey & KEY_SEP & label -> Double
    Set oRowData = CreateObject("Scripting.Dictionary")
    Set oColumns = CreateObject("Scripting.Dictionary")
    Set oPivot = CreateObject("Scripting.Dictionary")

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

        If bHasZones And Len(sZoneProp) > 0 Then
            AccumulateZoneLengths oEl, sCableKey, zones, sZoneProp, oColumns, oPivot
        End If
    Next i

    ' --- Write to Excel (always create the workbook, even with zero pivot columns) ---
    WriteToExcel oRowData, oColumns, oPivot, rowOrder, Filepath, ExcelVisible

    ShowStatus GetTranslation("CableReportComplete", nRowCount, oColumns.Count, nIncomplete)
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

' AccumulateZoneLengths
' For each zone whose partial length inside it (Length.GetPartialLengthInsideZones, 1-element
' array - the exact per-zone idiom ExportLengthInRegion.AggregateByZoneAndProperty already uses)
' is > 0, sums that length into oPivot keyed cable-major/label-minor (sCableKey & KEY_SEP & label) -
' unlike that sibling's zone-index-major key, so two zones sharing a label for the SAME cable
' collide into the SAME key and sum, by design. A zone with no/blank property value contributes
' nothing (silently - a data-quality gap in the drawing, not worth a status flicker per occurrence).
Private Sub AccumulateZoneLengths(ByVal oEl As Element, ByVal sCableKey As String, ByRef zones() As Element, _
                                  ByVal sZoneProp As String, ByRef oColumns As Object, ByRef oPivot As Object)
    On Error GoTo ErrorHandler

    If Not HasElements(zones) Then Exit Sub

    Dim oneZone(0 To 0) As Element
    Dim i               As Long
    Dim dLen            As Double
    Dim sLabel          As String
    Dim vVal            As Variant
    Dim sKey            As String

    For i = LBound(zones) To UBound(zones)
        Set oneZone(0) = zones(i)
        dLen = Length.GetPartialLengthInsideZones(oEl, oneZone)
        If dLen > 0 Then
            vVal = CustomPropertyHandler.GetPropertyValueFromElement(zones(i), sZoneProp, sZoneProp)
            sLabel = ""
            If Not IsNull(vVal) Then sLabel = Trim(CStr(vVal))
            If Len(sLabel) > 0 Then
                If Not oColumns.Exists(sLabel) Then oColumns.Add sLabel, True
                sKey = sCableKey & KEY_SEP & sLabel
                If oPivot.Exists(sKey) Then
                    oPivot(sKey) = oPivot(sKey) + dLen
                Else
                    oPivot.Add sKey, dLen
                End If
            End If
        End If
    Next i
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport.AccumulateZoneLengths"
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
' Late-binds Excel, writes fixed headers (Repere/Nature/Longueur) + one alphabetically-sorted
' pivot column per distinct soil-type label, one row per cable in scan order (rowOrder - no
' natural sort key for a Repere string). A pivot cell with no contribution stays BLANK, never 0.
'
' COM lifecycle: identical contract to ExportLengthInRegion.WriteToExcel (bExcelStartedByUs,
' always-attempted close on error, Quit gated on having started the session ourselves).
Private Sub WriteToExcel(ByRef oRowData As Object, ByRef oColumns As Object, ByRef oPivot As Object, _
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
    Dim sKey              As String
    Dim sPivotKey         As String
    Dim vRow              As Variant

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

    ' (3) Fixed headers + one sorted column per distinct soil-type label.
    xlSheet.Cells(1, 1).Value = GetTranslation("CableReportHeaderRepere")
    xlSheet.Cells(1, 2).Value = GetTranslation("CableReportHeaderNature")
    xlSheet.Cells(1, 3).Value = GetTranslation("CableReportHeaderLongueur")

    If oColumns.Count > 0 Then colKeys = SortedKeysCI(oColumns)
    For c = 0 To oColumns.Count - 1
        xlSheet.Cells(1, 4 + c).Value = colKeys(c)
    Next c

    ' (4) Data rows, one per cable, in scan order.
    For i = LBound(rowOrder) To UBound(rowOrder)
        sKey = rowOrder(i)
        r = i - LBound(rowOrder) + 2
        vRow = oRowData(sKey)
        xlSheet.Cells(r, 1).Value = vRow(0)
        If Not IsNull(vRow(1)) Then xlSheet.Cells(r, 2).Value = vRow(1)
        If Not IsNull(vRow(2)) Then xlSheet.Cells(r, 3).Value = vRow(2)
        For c = 0 To oColumns.Count - 1
            sPivotKey = sKey & KEY_SEP & colKeys(c)
            If oPivot.Exists(sPivotKey) Then xlSheet.Cells(r, 4 + c).Value = Round(oPivot(sPivotKey), nRound)
        Next c
    Next i

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
