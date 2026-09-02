' Module: PropertyCalculation_SourceEval
' Description: Source evaluation - cell/level/group scans, color/coord/length reads for every calc-rule
'              Source kind. Pure over its oEl/CalcRuleInfo/pattern arguments; no module-scope state.
'              Full grammar, engine passes and coupling doctrine: see _bmad/docs/calc-rules-grammar.md.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (global ARESConfig), RuleGrammar, PropertyRendering, Link,
'               Length, CustomPropertyHandler, ErrorHandlerClass (global ErrorHandler), PropertyCalculation
'               (Core - Types, HasElements, CascadeTrigger, ReportMultipleColorCandidates,
'               ReportMultiplePropCandidates, ReportMultipleGeometries, ReportNoGeoReference,
'               ReportUnknownGeoSystem),
'               ActiveModelReference/GeographicCoordinateSystem/Application.CreateGCSFromKeyName
'               (geo output - see calc-rules-grammar.md)

Option Explicit

' Runtime clamp so VBA Round never faults on an absurd decimal count (no coordinate/length ever needs > 15
' places). Relocated here from Core: this plan listed it among Core's "grammar consts", but its only two
' callers (FormatCoord, EvaluateOwnLength) both move to this module - found during implementation, same
' "travels with its consumer" pattern as PropertyRendering's VariantToPlainString/CountLiveEntries.
Private Const SOURCE_ROUND_CLAMP As Long = 15

' Evaluate a calc rule's Source against the bearing element. Returns the computed/fixed string ("" when a
' CellText/GroupLength source finds no matching member, or a SELF attribute is unavailable). Coordinates are
' ALREADY master units (mvba-docs) - no scaling.
Public Function EvaluateSource(ByRef r As CalcRuleInfo, ByVal oEl As element) As String
    On Error GoTo ErrorHandler

    EvaluateSource = ""
    Select Case r.SourceKind
        Case csCellText, csCellId, csCellLvl, csCellColor, csCellStyle, csCellWeight
            EvaluateSource = EvaluateGroupCellSource(oEl, r.SourceArg, r.SourceKind)
        Case csCellCoord
            EvaluateSource = EvaluateGroupCellSource(oEl, r.SourceArg, r.SourceKind, r.SourceSystem, r.SourceGeoDecimals)
        Case csLvlColor, csLvlStyle, csLvlWeight
            EvaluateSource = EvaluateGroupLvlSource(oEl, r.SourceArg, r.SourceKind)
        Case csValue
            EvaluateSource = r.SourceArg
        Case csCoord
            Dim pt As Point3d
            If GetElementAnchorPoint(oEl, pt) Then
                If Len(r.SourceSystem) > 0 Then
                    Dim sGeo As String
                    If ConvertToGeoSystem(pt, r.SourceSystem, r.SourceGeoDecimals, sGeo) Then
                        EvaluateSource = sGeo
                    Else
                        EvaluateSource = ""            ' one-shot status already reported by ConvertToGeoSystem
                    End If
                Else
                    EvaluateSource = FormatCoord(pt, ResolveDecimals(r.SourceArg, GetCoordDefaultDecimals()))
                End If
            Else
                ' No valid anchor (a geometry fault, or a non-graphical bearing element) -> yield NO value
                ' (never a fabricated "0;0"), the same "no value rather than a wrong value" philosophy as a
                ' FormatCoord fault. Log the technical anomaly (English, Number 0); ResolvePropertyValue then
                ' returns "" -> the transition-guarded clear (safe).
                ErrorHandler.HandleError "Property calculation: no anchor point for Coord source", 0, "", "PropertyCalculation_SourceEval.EvaluateSource"
                EvaluateSource = ""
            End If
        Case csId
            EvaluateSource = DLongToString(oEl.ID)
        Case csLvl
            If oEl.IsGraphical Then
                If Not oEl.Level Is Nothing Then EvaluateSource = oEl.Level.Name
            End If
        Case csColor
            If oEl.IsGraphical Then EvaluateSource = CStr(oEl.Color)
        Case csStyle
            If oEl.IsGraphical Then
                If Not oEl.LineStyle Is Nothing Then EvaluateSource = oEl.LineStyle.Name
            End If
        Case csWeight
            If oEl.IsGraphical Then EvaluateSource = CStr(oEl.LineWeight)
        Case csGroupColor
            EvaluateSource = EvaluateGroupColor(oEl)
        Case csGroupProp
            EvaluateSource = EvaluateGroupProp(oEl, r.SourceArg)
        Case csLength
            EvaluateSource = EvaluateOwnLength(oEl, ResolveDecimals(r.SourceArg, GetLengthDefaultDecimals()))
        Case csGroupLength
            EvaluateSource = EvaluateGroupLength(oEl, ResolveDecimals(r.SourceArg, GetLengthDefaultDecimals()))
    End Select
    Exit Function

ErrorHandler:
    EvaluateSource = ""
End Function

' n from a Coord[n]/Length[n]/GroupLength[n] SourceArg ("" -> defaultDec).
Private Function ResolveDecimals(ByVal sArg As String, ByVal defaultDec As Long) As Long
    If Len(sArg) > 0 Then
        ResolveDecimals = CLng(sArg)
    Else
        ResolveDecimals = defaultDec
    End If
End Function

' Shared GROUP scan for every Cell* source: scan the group INCLUDING itself, return the FIRST matching
' cell via foundCell and the total match count via nMatch (>= 2 drives the multi-trigger warning). An
' ungrouped bearing element is its own sole candidate.
Private Function FindFirstMatchingCellInGroup(ByVal oEl As element, ByVal sPattern As String, ByRef foundCell As element, ByRef nMatch As Long) As Boolean
    On Error GoTo ErrorHandler

    FindFirstMatchingCellInGroup = False
    Set foundCell = Nothing
    nMatch = 0

    Dim cands() As element
    cands = Link.GetLink(oEl, True)

    If PropertyCalculation.HasElements(cands) Then
        Dim i As Long
        For i = LBound(cands) To UBound(cands)
            If IsMatchingCell(cands(i), sPattern) Then
                nMatch = nMatch + 1
                If foundCell Is Nothing Then Set foundCell = cands(i)
            End If
        Next i
    Else
        ' Ungrouped bearing element: it is its own (single) candidate.
        If IsMatchingCell(oEl, sPattern) Then
            nMatch = 1
            Set foundCell = oEl
        End If
    End If

    FindFirstMatchingCellInGroup = Not (foundCell Is Nothing)
    Exit Function

ErrorHandler:
    FindFirstMatchingCellInGroup = False
    Set foundCell = Nothing
    nMatch = 0
End Function

' Every Cell* source evaluation, unified: find the FIRST group cell matching sPattern (self-included via
' FindFirstMatchingCellInGroup) and read the attribute named by kind off THAT cell (ReadCellSourceValue); ""
' when no cell matches. >= 2 matches -> the multi-trigger warning (one-shot) - the same ambiguity regardless
' of WHICH attribute is being read off the matching cell. sSystem/sGeoDecimals are CellCoord's own
' geo-output options (§2.3a of the WGS84 plan) - "" for every other kind, unused by ReadCellSourceValue in
' that case.
Private Function EvaluateGroupCellSource(ByVal oEl As element, ByVal sPattern As String, ByVal kind As CalcSource, Optional ByVal sSystem As String = "", Optional ByVal sGeoDecimals As String = "") As String
    On Error GoTo ErrorHandler

    EvaluateGroupCellSource = ""

    Dim foundCell As element
    Dim nMatch As Long
    If FindFirstMatchingCellInGroup(oEl, sPattern, foundCell, nMatch) Then
        EvaluateGroupCellSource = ReadCellSourceValue(foundCell, kind, sSystem, sGeoDecimals)
    End If
    If nMatch >= 2 Then PropertyCalculation.ReportMultipleTriggers
    Exit Function

ErrorHandler:
    EvaluateGroupCellSource = ""
End Function

' Read ONE attribute off a SPECIFIC cell element (already located - either by FindFirstMatchingCellInGroup
' during the bearing pass, or as the trigger cell itself during the push pass). Never fabricates a value:
' a missing Level/LineStyle yields "" (mirrors the no-anchor Coord/CellCoord philosophy). Native CellCoord
' (sSystem = "") uses the default decimals (no [n] override - the bracket already carries the pattern);
' sSystem/sGeoDecimals (CellCoord's own geo-output options, see calc-rules-grammar.md) override that for the
' geo-output case ONLY (sSystem non-empty) - every other kind ignores both params. Public: also called
' directly by Module C
' (TriggerPush), which MUST thread the same rule's SourceSystem/SourceGeoDecimals through here for CellCoord
' - the bearing pass and the trigger-cell push pass would otherwise silently disagree on output format for
' the same rule (see the WGS84 plan's §3.1).
Public Function ReadCellSourceValue(ByVal oCell As element, ByVal kind As CalcSource, Optional ByVal sSystem As String = "", Optional ByVal sGeoDecimals As String = "") As String
    On Error GoTo ErrorHandler

    ReadCellSourceValue = ""
    Select Case kind
        Case csCellText
            ' A sub-text the RENDERER writes must not feed the value that governs it (self-ratchet). The
            ' exclusion is a function of the SOURCE CELL, resolved here rather than threaded through
            ' every caller.
            Dim exIds() As Long
            Dim nEx As Long
            If PropertyRendering.GetExcludedSubIds(oCell, exIds, nEx) Then
                ReadCellSourceValue = StringsInEl.GetConcatenatedTextExcluding(oCell, exIds, nEx)
            Else
                ReadCellSourceValue = StringsInEl.GetConcatenatedText(oCell)
            End If
        Case csCellCoord
            Dim pt As Point3d
            If GetElementAnchorPoint(oCell, pt) Then
                If Len(sSystem) > 0 Then
                    Dim sGeo As String
                    If ConvertToGeoSystem(pt, sSystem, sGeoDecimals, sGeo) Then
                        ReadCellSourceValue = sGeo
                    End If
                    ' else: "" already set above; one-shot status already reported by ConvertToGeoSystem
                Else
                    ReadCellSourceValue = FormatCoord(pt, GetCoordDefaultDecimals())
                End If
            Else
                ErrorHandler.HandleError "Property calculation: no anchor point for CellCoord source", 0, "", "PropertyCalculation_SourceEval.ReadCellSourceValue"
            End If
        Case csCellId
            ReadCellSourceValue = DLongToString(oCell.ID)
        Case csCellLvl
            If Not oCell.Level Is Nothing Then ReadCellSourceValue = oCell.Level.Name
        Case csCellColor
            ' FillMode=2-aware resolution - see ReadLvlSourceValue's csLvlColor comment for the full
            ' rationale. CellColor has no ByLevel/ByCell symbolic state of its own, so
            ' ResolveFillAwareColor's result is used as-is.
            ReadCellSourceValue = CStr(ResolveFillAwareColor(oCell))
        Case csCellStyle
            If Not oCell.LineStyle Is Nothing Then ReadCellSourceValue = oCell.LineStyle.Name
        Case csCellWeight
            ReadCellSourceValue = CStr(oCell.LineWeight)
    End Select
    Exit Function

ErrorHandler:
    ReadCellSourceValue = ""
End Function

' True when sName matches at least one of sPattern's "|" -separated parts (case-insensitive, wildcards) -
' the same alternation grammar as a tag/calc CONDITION. Shared by every cell-name-vs-pattern comparison.
' Public: also called directly by Module C (TriggerPush)'s own IsMatchingCell/IsMatchingLevel-equivalent
' pattern checks.
Public Function MatchesAnyPattern(ByVal sName As String, ByVal sPattern As String) As Boolean
    On Error GoTo ErrorHandler

    MatchesAnyPattern = False

    Dim parts() As String
    parts = Split(sPattern, ARESConstants.ARES_VAR_DELIMITER)

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            If RuleGrammar.LikeCI(sName, parts(i)) Then
                MatchesAnyPattern = True
                Exit Function
            End If
        End If
    Next i
    Exit Function

ErrorHandler:
    MatchesAnyPattern = False
End Function

' True when el is a cell whose name matches sPattern (MatchesAnyPattern - wildcards + "|" alternation).
Private Function IsMatchingCell(ByVal el As element, ByVal sPattern As String) As Boolean
    On Error GoTo ErrorHandler

    IsMatchingCell = False
    If el Is Nothing Then Exit Function
    If Not el.IsCellElement Then Exit Function
    IsMatchingCell = MatchesAnyPattern(el.AsCellElement.Name, sPattern)
    Exit Function

ErrorHandler:
    IsMatchingCell = False
End Function

' True when el carries a Level whose name matches sPattern (MatchesAnyPattern - wildcards + "|"
' alternation). Unlike IsMatchingCell, NO element-type restriction: a Level can be carried by any graphical
' element (Line, Arc, Shape, cell, text...), not just a cell - the whole point of Lvl* sources is to cover
' the case where the group's authority is a plain geometry on a named level, not a named cell.
Private Function IsMatchingLevel(ByVal el As element, ByVal sPattern As String) As Boolean
    On Error GoTo ErrorHandler

    IsMatchingLevel = False
    If el Is Nothing Then Exit Function
    If Not el.IsGraphical Then Exit Function
    If el.Level Is Nothing Then Exit Function
    IsMatchingLevel = MatchesAnyPattern(el.Level.Name, sPattern)
    Exit Function

ErrorHandler:
    IsMatchingLevel = False
End Function

' Shared GROUP scan for every Lvl* source, mirroring FindFirstMatchingCellInGroup but matching on LEVEL
' name instead of cell name.
Private Function FindFirstMatchingLevelInGroup(ByVal oEl As element, ByVal sPattern As String, ByRef foundEl As element, ByRef nMatch As Long) As Boolean
    On Error GoTo ErrorHandler

    FindFirstMatchingLevelInGroup = False
    Set foundEl = Nothing
    nMatch = 0

    Dim cands() As element
    cands = Link.GetLink(oEl, True)

    If PropertyCalculation.HasElements(cands) Then
        Dim i As Long
        For i = LBound(cands) To UBound(cands)
            If IsMatchingLevel(cands(i), sPattern) Then
                nMatch = nMatch + 1
                If foundEl Is Nothing Then Set foundEl = cands(i)
            End If
        Next i
    Else
        ' Ungrouped bearing element: it is its own (single) candidate.
        If IsMatchingLevel(oEl, sPattern) Then
            nMatch = 1
            Set foundEl = oEl
        End If
    End If

    FindFirstMatchingLevelInGroup = Not (foundEl Is Nothing)
    Exit Function

ErrorHandler:
    FindFirstMatchingLevelInGroup = False
    Set foundEl = Nothing
    nMatch = 0
End Function

' Every Lvl* source evaluation, unified: read kind off the first Level-matching group member. Uses its
' OWN ReportMultipleLvlTriggers, not ReportMultipleTriggers - a Lvl* collision may involve no cell at all.
Private Function EvaluateGroupLvlSource(ByVal oEl As element, ByVal sPattern As String, ByVal kind As CalcSource) As String
    On Error GoTo ErrorHandler

    EvaluateGroupLvlSource = ""

    Dim foundEl As element
    Dim nMatch As Long
    If FindFirstMatchingLevelInGroup(oEl, sPattern, foundEl, nMatch) Then
        EvaluateGroupLvlSource = ReadLvlSourceValue(foundEl, kind)
    End If
    If nMatch >= 2 Then PropertyCalculation.ReportMultipleLvlTriggers
    Exit Function

ErrorHandler:
    EvaluateGroupLvlSource = ""
End Function

' csLvlColor/csLvlWeight: Color/LineWeight is a 3-state MicroStation value (explicit / ByLevel / ByCell).
' Only ByLevel resolves here (via the element's own Level default); ByCell yields "" - never fabricate.
' csLvlColor tests ResolveFillAwareColor's result, not .Color directly, so FillMode=2 is resolved first.
' csLvlStyle has no ByLevel/ByCell equivalent (LineStyle is an object, not a sentinel) - reads it as-is.
' Public: also called directly by Module C (TriggerPush).
Public Function ReadLvlSourceValue(ByVal oFoundEl As element, ByVal kind As CalcSource) As String
    On Error GoTo ErrorHandler

    ReadLvlSourceValue = ""
    Select Case kind
        Case csLvlColor
            Dim rawLvlColor As Long
            rawLvlColor = ResolveFillAwareColor(oFoundEl)
            If rawLvlColor = ByLevelColor Then
                If Not oFoundEl.Level Is Nothing Then ReadLvlSourceValue = CStr(oFoundEl.Level.ElementColor)
            ElseIf rawLvlColor = ByCellColor Then
                ReadLvlSourceValue = ""                ' not resolvable from the Level - never fabricate
            Else
                ReadLvlSourceValue = CStr(rawLvlColor)
            End If
        Case csLvlStyle
            If Not oFoundEl.LineStyle Is Nothing Then ReadLvlSourceValue = oFoundEl.LineStyle.Name
        Case csLvlWeight
            If oFoundEl.LineWeight = ByLevelLineWeight Then
                If Not oFoundEl.Level Is Nothing Then ReadLvlSourceValue = CStr(oFoundEl.Level.ElementLineWeight)
            ElseIf oFoundEl.LineWeight = ByCellLineWeight Then
                ReadLvlSourceValue = ""                ' not resolvable from the Level - never fabricate
            Else
                ReadLvlSourceValue = CStr(oFoundEl.LineWeight)
            End If
    End Select
    Exit Function

ErrorHandler:
    ReadLvlSourceValue = ""
End Function

' Shared FillMode=2-aware color resolution (GroupColor/CellColor/LvlColor): a ClosedElement in FillMode=2
' reads its FILL color unless that fill is literally 0/255, falling back to .Color. Callers that care
' about the ByLevel/ByCell sentinel test THIS function's result, not el.Color directly.
Private Function ResolveFillAwareColor(ByVal el As element) As Long
    On Error GoTo ErrorHandler

    ResolveFillAwareColor = el.Color
    If el.IsClosedElement Then
        If el.AsClosedElement.FillMode = 2 Then
            Dim fc As Long
            fc = el.AsClosedElement.fillcolor
            If fc = 0 Or fc = 255 Then
                ResolveFillAwareColor = el.Color
            Else
                ResolveFillAwareColor = fc
            End If
        End If
    End If
    Exit Function

ErrorHandler:
    ResolveFillAwareColor = el.Color
End Function

' FreshHandle - the rule every PULL source (GroupProp/GroupColor/GroupLength) follows when reading a scanned
' group member: if that member IS this cascade's trigger, read it through the trigger's OWN handle
' (PropertyCalculation.CascadeTrigger), not the one the scan returned - a scan run inside an element's own
' ElementChanged callback still sees its PRE-change state, one edit behind (see mvba-cheatsheet.md). The scan
' still decides membership; only the READ moves. Falls back to the scanned handle when there is no cascade
' trigger, or when the trigger handle is no longer readable (deleted mid-cascade - .ID raises).
Private Function FreshHandle(ByVal oCand As element) As element
    On Error GoTo Fallback

    Set FreshHandle = oCand
    If oCand Is Nothing Then Exit Function

    Dim oTrig As element
    Set oTrig = PropertyCalculation.CascadeTrigger
    If oTrig Is Nothing Then Exit Function

    If DLongComp(oCand.ID, oTrig.ID) = 0 Then Set FreshHandle = oTrig
    Exit Function

Fallback:
    Set FreshHandle = oCand
End Function

' GroupColor: unlike every other GROUP source, SELF-EXCLUDED (no ReturnMe:=True) - Color has no type
' filter to protect a self-match the way GroupLength's IsLengthCapableType does, so self-inclusion would
' make a bearing text match itself instead of its linked geometry. No name/level pattern - first linked
' element wins. >= 2 candidates -> CalculationMultipleColorCandidates (discloses the pick, doesn't change it).
Private Function EvaluateGroupColor(ByVal oEl As element) As String
    On Error GoTo ErrorHandler

    EvaluateGroupColor = ""

    Dim cands() As element
    cands = Link.GetLink(oEl)                      ' NO ReturnMe:=True - see rationale above
    If Not PropertyCalculation.HasElements(cands) Then Exit Function

    Dim candidate As element
    Set candidate = FreshHandle(cands(LBound(cands)))   ' pull source: read the trigger off its fresh handle
    If Not candidate Is Nothing Then
        EvaluateGroupColor = CStr(ResolveFillAwareColor(candidate))
    End If

    If UBound(cands) > LBound(cands) Then PropertyCalculation.ReportMultipleColorCandidates
    Exit Function

ErrorHandler:
    EvaluateGroupColor = ""
End Function

' GroupProp[PropertyName]: self-EXCLUDED (no ReturnMe:=True) like GroupColor - a bearer carrying its own copy
' of PropertyName must never satisfy its own search - and, unlike GroupColor, filtered on ATTACH STATE AND on
' a genuinely NON-EMPTY value: an attached-but-empty candidate is an unresolved RECEIVER of this same rule
' (a text just fan-out-attached), never a valid donor, and letting its blank win would overwrite a real value
' elsewhere. "" when ungrouped, uncarried, or every carrier is empty; >= 2 real-valued candidates raise
' CalculationMultiplePropCandidates, first non-empty in scan order winning. The donor/receiver
' rule-conditioning doctrine this filter complements: see calc-rules-grammar.md's GroupProp bullet.
Private Function EvaluateGroupProp(ByVal oEl As element, ByVal sPropName As String) As String
    On Error GoTo ErrorHandler

    EvaluateGroupProp = ""
    If Len(sPropName) = 0 Then Exit Function

    Dim cands() As element
    cands = Link.GetLink(oEl)                      ' NO ReturnMe:=True - see rationale above
    If Not PropertyCalculation.HasElements(cands) Then Exit Function

    Dim i        As Long
    Dim nFound   As Long
    Dim vVal     As Variant
    Dim sVal     As String
    Dim oRead    As element
    nFound = 0
    For i = LBound(cands) To UBound(cands)
        Set oRead = FreshHandle(cands(i))          ' pull source: the scan found it, the fresh handle is read
        If CustomPropertyHandler.IsItemAttachedToElement(oRead, sPropName) Then
            vVal = CustomPropertyHandler.GetPropertyValueFromElement(oRead, sPropName, sPropName)
            If Not IsNull(vVal) Then
                sVal = CStr(vVal)
                If Len(sVal) > 0 Then
                    nFound = nFound + 1
                    If nFound = 1 Then EvaluateGroupProp = sVal
                End If
            End If
        End If
    Next i

    If nFound > 1 Then PropertyCalculation.ReportMultiplePropCandidates
    Exit Function

ErrorHandler:
    EvaluateGroupProp = ""
End Function

'######################################################################################################################
'                                          GEOMETRY - Coord ANCHOR CASCADE
'######################################################################################################################

' Deterministic anchor point for Coord/CellCoord. The Range centre is seeded FIRST and a type-specific
' anchor overrides it only on success, so a per-branch geometry fault degrades to the Range centre, never
' to a fabricated (0,0,0). False only when even the Range seed fails.
Private Function GetElementAnchorPoint(ByVal oEl As element, ByRef pt As Point3d) As Boolean
    On Error GoTo ErrorHandler

    GetElementAnchorPoint = False
    If oEl Is Nothing Then Exit Function
    If Not oEl.IsGraphical Then Exit Function

    ' Seed = the universal Range centre. If even this faults (Range raises when not graphical - gated above),
    ' we have NO anchor and return False.
    Dim rng As Range3d
    rng = oEl.Range
    pt = Point3dFromXYZ((rng.Low.X + rng.High.X) / 2#, _
                        (rng.Low.Y + rng.High.Y) / 2#, _
                        (rng.Low.Z + rng.High.Z) / 2#)
    GetElementAnchorPoint = True                  ' at least the Range-centre seed is valid

    ' A type-specific anchor overrides the seed ONLY on success; a fault leaves the seed standing.
    Dim ptSpecific As Point3d
    If TryGetSpecificAnchor(oEl, ptSpecific) Then pt = ptSpecific
    Exit Function

ErrorHandler:
    ' Even the Range seed faulted -> no valid anchor (the caller yields "" + logs), never a fabricated coord.
    GetElementAnchorPoint = False
End Function

' Type-specific anchor per element type (cell/text/line Origin, arc/ellipse CenterPoint, closed shape
' Centroid - undocumented for closed elements, hence isolated here). False (-> caller keeps the Range
' seed) when the element has no specific anchor or the read faults.
Private Function TryGetSpecificAnchor(ByVal oEl As element, ByRef ptOut As Point3d) As Boolean
    On Error GoTo ErrorHandler

    Select Case True
        Case oEl.IsCellElement
            ptOut = oEl.AsCellElement.Origin
        Case oEl.Type = msdElementTypeSharedCell
            ptOut = oEl.AsSharedCellElement.Origin
        Case oEl.IsTextElement
            ptOut = oEl.AsTextElement.Origin
        Case oEl.IsTextNodeElement
            ptOut = oEl.AsTextNodeElement.Origin
        Case oEl.Type = msdElementTypeLine
            ptOut = oEl.AsLineElement.Origin
        Case oEl.Type = msdElementTypeArc
            ptOut = oEl.AsArcElement.CenterPoint
        Case oEl.Type = msdElementTypeEllipse
            ptOut = oEl.AsEllipseElement.CenterPoint
        Case oEl.IsClosedElement
            ptOut = oEl.AsClosedElement.Centroid
        Case Else
            TryGetSpecificAnchor = False
            Exit Function                          ' no specific anchor -> the caller uses the Range seed
    End Select

    TryGetSpecificAnchor = True
    Exit Function

ErrorHandler:
    ' Silent (no log): a specific-anchor read faulted -> no specific anchor; the caller keeps the Range-
    ' centre seed. The anomaly is logged once by EvaluateSource only if even the seed is unavailable.
    TryGetSpecificAnchor = False
End Function

' CStr/Format are locale-dependent (a comma-decimal Windows locale, e.g. fr-FR, turns 1.5 into "1,5") -
' VBA's Str() is documented to always use "." regardless of locale, at the cost of a leading space for a
' non-negative number, hence Trim. Asketyll wants coordinates (native and geographic) always "." - decided
' 2026-08-27, deliberately NOT applied to Length/GroupLength (those stay locale-native, unaffected).
Private Function FormatDotDecimal(ByVal d As Double) As String
    FormatDotDecimal = Trim(Str(d))
End Function

' Format a point as "X;Y" with dec decimals, mirroring Auto_Lengths.cls CStr(Round(...)) except the
' decimal separator is forced to "." (FormatDotDecimal) regardless of locale - the ";" field separator
' does not collide either way. dec is clamped so VBA Round never faults on an absurd count.
Private Function FormatCoord(ByRef pt As Point3d, ByVal dec As Long) As String
    On Error GoTo ErrorHandler

    Dim d As Long
    d = dec
    If d < 0 Then d = 0
    If d > SOURCE_ROUND_CLAMP Then d = SOURCE_ROUND_CLAMP

    FormatCoord = FormatDotDecimal(Round(pt.X, d)) & ";" & FormatDotDecimal(Round(pt.Y, d))
    Exit Function

ErrorHandler:
    FormatCoord = ""
End Function

'######################################################################################################################
'                                          GEOMETRY - GEOGRAPHIC OUTPUT (Coord/CellCoord[...][system,n])
'######################################################################################################################

' Converts a resolved DGN point to lat/long in the REQUESTED output system (sSystemName - "WGS84",
' "EPSG:4171", any name MicroStation's GeographicCoordinateSystem library recognises; ARES does not
' enumerate them, see calc-rules-grammar.md), and formats it "Latitude;Longitude" (confirmed field order).
' Two DISTINCT failure causes, each with its OWN one-shot status (reported HERE, not by the caller, since
' this is the only code that knows which of the two calls failed):
'  1. ActiveModelReference.GetGCS(True) returns Nothing - the ACTIVE MODEL has no georeferencing at all, so
'     there is no source datum to convert FROM (ReportNoGeoReference).
'  2. Application.CreateGCSFromKeyName(sSystemName) does not resolve - the REQUESTED system name is not
'     recognised (ReportUnknownGeoSystem) - a typo'd/unknown rule, not a model-configuration problem.
' NEVER attempts to configure a GCS itself in either case: both GetGCS(True) and CreateGCSFromKeyName are
' plain lookups (verified against mvba-docs - the official GeographicCoordinateSystem example uses
' CreateGCSFromKeyName with no interactive side effect; its own On Error Resume Next around that call is
' mirrored below, since the docs do not state whether an unresolved key name returns Nothing or raises).
' ARES has NO role in configuring a GCS - that is MicroStation's own GEOCOORDINATE SELECT LIBRARY key-in,
' run manually by the user; this module must NEVER call CadInputQueue.SendCommand or any other interactive
' API from this path (decided, not a coding-phase detail - see the WGS84 plan's §1/§1.1/§1.2). sGeoDecimals
' "" means FULL PRECISION (no Round call) - deliberately NOT the same "omitted = default rounding"
' convention as Coord[n]/Length[n]/GroupLength[n], since degrees and DGN master units have unrelated
' precision needs (decided by the lead, not guessed).
Private Function ConvertToGeoSystem(ByRef pt As Point3d, ByVal sSystemName As String, ByVal sGeoDecimals As String, ByRef sOut As String) As Boolean
    On Error GoTo ErrorHandler

    ConvertToGeoSystem = False
    sOut = ""

    Dim gcs As GeographicCoordinateSystem
    Set gcs = ActiveModelReference.GetGCS(True)
    If gcs Is Nothing Then
        PropertyCalculation.ReportNoGeoReference
        Exit Function                                ' no value - fail closed, no UI (see header comment)
    End If

    Dim otherGcs As GeographicCoordinateSystem
    On Error Resume Next
    Set otherGcs = Application.CreateGCSFromKeyName(sSystemName)
    On Error GoTo ErrorHandler
    If otherGcs Is Nothing Then
        PropertyCalculation.ReportUnknownGeoSystem
        Exit Function                                ' unrecognised system name - fail closed, no UI
    End If

    Dim p2d As Point2d
    p2d.X = pt.X
    p2d.y = pt.y

    Dim geo As GeoPoint2D
    geo = gcs.LatLongFromMasterUnits2d(p2d)          ' active model's own datum first
    geo = gcs.LatLongFromLatLong2d(geo, otherGcs)     ' then reprojected to the REQUESTED system's datum

    sOut = FormatGeoPoint(geo, sGeoDecimals)
    ConvertToGeoSystem = True
    Exit Function

ErrorHandler:
    ConvertToGeoSystem = False
    sOut = ""
End Function

' Formats a GeoPoint2D as "Latitude;Longitude" - NOT FormatCoord's "X;Y": a different domain (degrees, not
' master units), a different default (full precision when sDec is "", vs FormatCoord's always-rounded
' default), and a confirmed field order that is not the same axis order as pt.X;pt.Y. Explicit digits (sDec
' non-empty) round exactly like FormatCoord, same SOURCE_ROUND_CLAMP-guarded Round. Decimal separator forced
' to "." via FormatDotDecimal, same as FormatCoord (2026-08-27) - locale-independent regardless of digit count.
Private Function FormatGeoPoint(ByRef geo As GeoPoint2D, ByVal sDec As String) As String
    On Error GoTo ErrorHandler

    If Len(sDec) = 0 Then
        FormatGeoPoint = FormatDotDecimal(geo.Latitude) & ";" & FormatDotDecimal(geo.Longitude)
    Else
        Dim d As Long
        d = CLng(sDec)
        If d < 0 Then d = 0
        If d > SOURCE_ROUND_CLAMP Then d = SOURCE_ROUND_CLAMP
        FormatGeoPoint = FormatDotDecimal(Round(geo.Latitude, d)) & ";" & FormatDotDecimal(Round(geo.Longitude, d))
    End If
    Exit Function

ErrorHandler:
    FormatGeoPoint = ""
End Function

'######################################################################################################################
'                                          GEOMETRY - LENGTH SOURCES (Length / GroupLength)
'######################################################################################################################

' True when oEl is one of the types Length.GetLength actually measures (Line/Arc/Shape/ComplexShape/
' ComplexString - the same set Auto_Lengths.GetLinkedElements filters for). A cell/text/other element is
' NOT length-capable - never call Length.GetLength on it (it would yield 0 + a noisy status line).
Private Function IsLengthCapableType(ByVal oEl As element) As Boolean
    On Error GoTo ErrorHandler

    IsLengthCapableType = False
    If oEl Is Nothing Then Exit Function
    IsLengthCapableType = oEl.IsLineElement Or oEl.IsArcElement Or oEl.IsShapeElement Or _
                           oEl.IsComplexShapeElement Or oEl.IsComplexStringElement
    Exit Function

ErrorHandler:
    IsLengthCapableType = False
End Function

' Length evaluation: the bearing element's OWN geometry length (Length.GetLength), ONLY when it is itself
' length-capable; "" otherwise (never a fabricated 0 - the same "no value rather than a wrong value"
' philosophy as Coord). dec is clamped so VBA Round never faults on an absurd count (mirrors FormatCoord).
Private Function EvaluateOwnLength(ByVal oEl As element, ByVal dec As Long) As String
    On Error GoTo ErrorHandler

    EvaluateOwnLength = ""
    If oEl Is Nothing Then Exit Function
    If Not IsLengthCapableType(oEl) Then Exit Function

    Dim d As Long
    d = dec
    If d < 0 Then d = 0
    If d > SOURCE_ROUND_CLAMP Then d = SOURCE_ROUND_CLAMP

    EvaluateOwnLength = CStr(Length.GetLength(oEl, CByte(d)))
    Exit Function

ErrorHandler:
    EvaluateOwnLength = ""
End Function

' Shared GROUP scan for GroupLength: first length-capable member by scan order (no name pattern -
' geometry has none, unlike Cell*). Mirrors FindFirstMatchingCellInGroup.
Private Function FindFirstLengthCapableInGroup(ByVal oEl As element, ByRef foundGeo As element, ByRef nMatch As Long) As Boolean
    On Error GoTo ErrorHandler

    FindFirstLengthCapableInGroup = False
    Set foundGeo = Nothing
    nMatch = 0

    Dim cands() As element
    cands = Link.GetLink(oEl, True)

    If PropertyCalculation.HasElements(cands) Then
        Dim i As Long
        For i = LBound(cands) To UBound(cands)
            If IsLengthCapableType(cands(i)) Then
                nMatch = nMatch + 1
                If foundGeo Is Nothing Then Set foundGeo = cands(i)
            End If
        Next i
    Else
        If IsLengthCapableType(oEl) Then
            nMatch = 1
            Set foundGeo = oEl
        End If
    End If

    FindFirstLengthCapableInGroup = Not (foundGeo Is Nothing)
    Exit Function

ErrorHandler:
    FindFirstLengthCapableInGroup = False
    Set foundGeo = Nothing
    nMatch = 0
End Function

' GroupLength: length of the FIRST length-capable member found (self-included scan order); "" when none.
' WARNS on multiple candidates - GroupLength replaced Auto Lengths, which refused to choose silently when
' several geometries were measurable, so this makes the same arbitrary pick visible rather than silent.
Private Function EvaluateGroupLength(ByVal oEl As element, ByVal dec As Long) As String
    On Error GoTo ErrorHandler

    EvaluateGroupLength = ""

    Dim foundGeo As element
    Dim nMatch As Long
    If FindFirstLengthCapableInGroup(oEl, foundGeo, nMatch) Then
        EvaluateGroupLength = EvaluateOwnLength(FreshHandle(foundGeo), dec)   ' pull source: fresh trigger handle
        If nMatch >= 2 Then PropertyCalculation.ReportMultipleGeometries
    End If
    Exit Function

ErrorHandler:
    EvaluateGroupLength = ""
End Function

' Default decimals for a bare Length/GroupLength source = the Auto Lengths rounding convention
' ARES_Length_Round (distinct from ARES_Round, used by Coord). Fail-closed to 2 on any nil; lazy ARESConfig
' init like the other readers.
Private Function GetLengthDefaultDecimals() As Long
    On Error GoTo ErrorHandler

    GetLengthDefaultDecimals = 2
    If ARESConfig Is Nothing Then Exit Function
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    If ARESConfig.ARES_LENGTH_ROUND Is Nothing Then Exit Function
    GetLengthDefaultDecimals = CLng(ARESConfig.ARES_LENGTH_ROUND.Value)
    Exit Function

ErrorHandler:
    GetLengthDefaultDecimals = 2
End Function

' Default decimals for a bare Coord source = the existing rounding convention ARES_Round (2). Fail-closed
' to 2 on any nil; lazy ARESConfig init like the other readers.
Private Function GetCoordDefaultDecimals() As Long
    On Error GoTo ErrorHandler

    GetCoordDefaultDecimals = 2
    If ARESConfig Is Nothing Then Exit Function
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    If ARESConfig.ARES_ROUNDS Is Nothing Then Exit Function
    GetCoordDefaultDecimals = CLng(ARESConfig.ARES_ROUNDS.Value)
    Exit Function

ErrorHandler:
    GetCoordDefaultDecimals = 2
End Function
