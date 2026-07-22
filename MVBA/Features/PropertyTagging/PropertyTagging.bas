' Module: PropertyTagging
' Description: The SOLE attach/detach engine for ARES custom properties. Auto-attaches ARES custom
'              properties to elements as they are created / modified, driven by configurable rules
'              (ARES_Property_Rules). Attach-only on match: the property appears (empty) on the element
'              for the user to fill / pick from its native value-list dropdown. The companion value
'              engine (PropertyPropagation) writes VALUES into these already-attached properties but
'              never attaches/detaches itself - it delegates any detach back here (DetachRuleProperty).
'
'              Called from ElementChangeHandler.ProcessElement (deferred, on idle) when
'              ARES_Auto_Properties is True. Rules are parsed once and cached - call RefreshRules
'              after changing ARES_Property_Rules at runtime.
'
'              Rule format (ARES_Property_Rules):  "selector=prop|prop ; selector=prop ; ..."  where a
'              selector is either a level rule or a cell-group rule:
'                - level[:type]  : MicroStation level name (required; case-insensitive) + optional
'                                  element type name (StringToMsdElementType; absent = any type). The
'                                  properties attach to the matching element ITSELF.
'                - @CellName     : cell-group rule (leading "@" marker). When a cell named CellName is
'                                  processed (add/modify, Depth 0) and sits in a real graphic group, the
'                                  properties attach to each OTHER member of that group (Link.GetLink) -
'                                  the same fan-out the value engine uses. Case-insensitive, trimmed.
'                - props : | -delimited property names (must exist in the "ARES" DGNLib)
'              Example:  WALLS=Commune|Coupe_Type ; DOORS:Cell=Commune ; @ETI076=Commune
'
'              DetachRuleProperty(El, P) is the public detach service (round-4): it removes a single
'              property, called by the value engine when it empties a value with the detach-empty
'              option ON. Keeping detach here preserves "only the tagger attaches/detaches".
'
'              GetCellGroupProperties(sCellName) is the read-only accessor the value engine
'              (PropertyPropagation) reads to derive its triggers and targets from the @cell rules -
'              the @cell rules are the SINGLE source (epic 12): a rule @X=P both attaches P to X's
'              group and marks X's text as the value written into P.
'
'              ValidateRuleSyntax(sRule) is the read-only grammar validator (epic 12-2) the options form
'              calls on every rule commit. It rejects the malformations the parser swallows silently
'              (chiefly a property token containing "=", "@" or ";" - the "|"-instead-of-";" mistake).
'              Syntactic only (no DGNLib membership check); reuses the module's grammar constants.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (global ARESConfig), CustomPropertyHandler, Link, MicroStationDefinition, ErrorHandlerClass (global ErrorHandler)

Option Explicit

Private Const RULE_SEPARATOR As String = ";"
Private Const SELECTOR_SEPARATOR As String = "="
Private Const TYPE_SEPARATOR As String = ":"
Private Const CELL_GROUP_MARKER As String = "@"

' One parsed rule. Either a LEVEL rule (Level [+ optional ElType] -> attach Props to the element
' itself) or a CELL-GROUP rule (IsCellGroup=True, CellName set -> attach Props to the OTHER members of
' the named cell's graphic group). Props is populated identically for both selector kinds.
Private Type RuleInfo
    Level As String
    HasType As Boolean
    ElType As MsdElementType
    IsCellGroup As Boolean
    CellName As String
    Props() As String
End Type

Private mRules() As RuleInfo
Private mnRuleCount As Long
Private mbParsed As Boolean

' Force a re-parse of ARES_Property_Rules on the next match/apply (call after editing the variable).
Public Sub RefreshRules()
    mbParsed = False
End Sub

' True when the element's level (and type, if the rule specifies one) matches at least one rule.
' Fast path for ElementChangeHandler.ShouldQueueElement.
Public Function ElementMatchesAnyRule(ByVal oElement As element) As Boolean
    On Error GoTo ErrorHandler

    ElementMatchesAnyRule = False
    If oElement Is Nothing Then Exit Function

    EnsureRulesParsed
    If mnRuleCount = 0 Then Exit Function

    Dim sLevel As String
    Dim i As Long
    ' Level-based rules: an element with no level matches nothing. Guard with nested Ifs (never And -
    ' no short-circuit in VBA): Level GET raises on a non-graphical element and can be Nothing (e.g. a
    ' cell header), which would make .Name raise Error 91. Silent no-match, no log.
    If Not oElement.IsGraphical Then Exit Function
    If oElement.Level Is Nothing Then Exit Function
    sLevel = oElement.Level.Name

    ' Cell-group rules are intentionally ignored here: they require a graphic group, so an ungrouped
    ' element (this fast-path's only caller, ShouldQueueElement) can never match one. Level rules only.
    For i = 0 To mnRuleCount - 1
        If Not mRules(i).IsCellGroup Then
            If RuleMatches(mRules(i), oElement, sLevel) Then
                ElementMatchesAnyRule = True
                Exit Function
            End If
        End If
    Next i
    Exit Function

ErrorHandler:
    ElementMatchesAnyRule = False
End Function

' Attach the configured properties for every rule the element drives. Two passes: (1) cell-group rules
' fan out from a named cell to its OTHER graphic-group members; (2) level[:type] rules attach to the
' element itself. All attaches are idempotent (CustomPropertyHandler.AttachItemToElement only attaches
' when not already present) -> loop-safe.
Public Sub ApplyPropertyRules(ByVal oElement As element)
    On Error GoTo ErrorHandler

    If oElement Is Nothing Then Exit Sub

    EnsureRulesParsed
    If mnRuleCount = 0 Then Exit Sub

    ' Cell-group pass FIRST, and independently of the level guards below: an IsCellElement element is
    ' graphical, but structuring the cell path ahead of the IsGraphical/Level guards keeps it reachable
    ' even for an element whose Level read would exit the level path. AttachCellGroupRules carries its
    ' own graphic-group guard. (IsCellElement is safe to read on any non-Nothing element - mirrors
    ' IsTriggerCell.)
    If oElement.IsCellElement Then AttachCellGroupRules oElement

    ' Level pass: level[:type] rules attach to the element itself. An element with no level matches
    ' nothing. Guard with nested Ifs (never And - no short-circuit in VBA): Level GET raises on a
    ' non-graphical element and can be Nothing (e.g. a cell header), which would make .Name raise
    ' Error 91. Silent skip, no log.
    If Not oElement.IsGraphical Then Exit Sub
    If oElement.Level Is Nothing Then Exit Sub

    Dim sLevel As String
    Dim i As Long, j As Long
    sLevel = oElement.Level.Name

    For i = 0 To mnRuleCount - 1
        If Not mRules(i).IsCellGroup Then
            If RuleMatches(mRules(i), oElement, sLevel) Then
                For j = LBound(mRules(i).Props) To UBound(mRules(i).Props)
                    If Len(mRules(i).Props(j)) > 0 Then
                        CustomPropertyHandler.AttachItemToElement oElement, mRules(i).Props(j)
                    End If
                Next j
            End If
        End If
    Next i
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging.ApplyPropertyRules"
End Sub

' Cell-group attach fan-out: when oCell (a CellElement) sits in a real graphic group and its name
' matches a cell-group rule (@CellName), attach that rule's properties to each OTHER member of the
' group (Link.GetLink excludes oCell). Attach is idempotent (AttachItemToElement is HasItems-guarded),
' so re-processing the same cell attaches nothing new -> loop-safe. Late-joining siblings are not
' auto-tagged until the cell is next touched (same limitation as the value engine's fan-out).
Private Sub AttachCellGroupRules(ByVal oCell As element)
    On Error GoTo ErrorHandler

    If oCell Is Nothing Then Exit Sub
    If oCell.GraphicGroup = ARES_DEFAULT_GRAPHIC_GROUP_ID Then Exit Sub

    Dim sName As String
    sName = oCell.AsCellElement.Name

    Dim els() As element
    Dim bResolved As Boolean
    bResolved = False

    Dim i As Long, j As Long, k As Long
    Dim s As element
    For i = 0 To mnRuleCount - 1
        If mRules(i).IsCellGroup Then
            If StrComp(Trim(mRules(i).CellName), Trim(sName), vbTextCompare) = 0 Then
                ' Resolve the group's OTHER members once (first matching rule) and reuse.
                If Not bResolved Then
                    els = Link.GetLink(oCell)
                    bResolved = True
                End If
                If HasElements(els) Then
                    For j = LBound(els) To UBound(els)
                        Set s = els(j)
                        If Not s Is Nothing Then
                            For k = LBound(mRules(i).Props) To UBound(mRules(i).Props)
                                If Len(mRules(i).Props(k)) > 0 Then
                                    CustomPropertyHandler.AttachItemToElement s, mRules(i).Props(k)
                                End If
                            Next k
                        End If
                    Next j
                End If
            End If
        End If
    Next i
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging.AttachCellGroupRules"
End Sub

' Public detach service (round-4): remove a single property P from El. Called by the value engine
' (PropertyPropagation) when it empties a value with ARES_Propagation_Detach_Empty ON - detach is
' delegated here so ALL attach/detach stays inside PropertyTagging. Thin wrapper over
' CustomPropertyHandler.RemoveItemFromElement (itself HasItems-guarded, idempotent). Does NOT consult
' the parsed rules: a governing rule re-attaches P (empty) on the next Depth-0 pass - the intended,
' terminating interaction (the value engine's non-empty->empty transition guard bounds it to one
' detach per emptying).
Public Sub DetachRuleProperty(ByVal El As element, ByVal P As String)
    On Error GoTo ErrorHandler

    If El Is Nothing Then Exit Sub
    If Len(Trim(P)) = 0 Then Exit Sub

    CustomPropertyHandler.RemoveItemFromElement El, P
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging.DetachRuleProperty"
End Sub

' Read-only accessor for the value engine (PropertyPropagation): the aggregated non-empty properties of
' every @cell rule whose CellName matches sCellName (case-insensitive, trimmed). This is the SINGLE
' source the value engine uses to derive BOTH its triggers (>=1 property returned => the cell is a
' trigger) and its targets (the returned properties). Union across matching rules (@X=P1 ; @X=P2 ->
' {P1,P2}; @X=P1|P2 -> {P1,P2}), deduplicated case-insensitively. Returns a single-element [""] array
' (the SplitTrim empty convention) when the name is empty or no @cell rule matches, so callers test
' Len(result(LBound)) > 0. Read-only (parses the cache via EnsureRulesParsed; no model write, no attach/
' detach). Standard error pattern -> [""] on fault.
Public Function GetCellGroupProperties(ByVal sCellName As String) As String()
    On Error GoTo ErrorHandler

    Dim out() As String
    Dim n As Long
    ReDim out(0 To 0)
    out(0) = ""
    n = 0

    Dim sTarget As String
    sTarget = Trim(sCellName)
    If Len(sTarget) = 0 Then
        GetCellGroupProperties = out
        Exit Function
    End If

    EnsureRulesParsed
    If mnRuleCount = 0 Then
        GetCellGroupProperties = out
        Exit Function
    End If

    Dim i As Long, j As Long, m As Long
    Dim sProp As String
    Dim bDup As Boolean
    For i = 0 To mnRuleCount - 1
        If mRules(i).IsCellGroup Then
            If StrComp(Trim(mRules(i).CellName), sTarget, vbTextCompare) = 0 Then
                For j = LBound(mRules(i).Props) To UBound(mRules(i).Props)
                    sProp = mRules(i).Props(j)
                    If Len(sProp) > 0 Then
                        bDup = False
                        For m = 0 To n - 1
                            If StrComp(out(m), sProp, vbTextCompare) = 0 Then
                                bDup = True
                                Exit For
                            End If
                        Next m
                        If Not bDup Then
                            If n = 0 Then
                                out(0) = sProp
                            Else
                                ReDim Preserve out(0 To n)
                                out(n) = sProp
                            End If
                            n = n + 1
                        End If
                    End If
                Next j
            End If
        End If
    Next i

    GetCellGroupProperties = out
    Exit Function

ErrorHandler:
    ReDim out(0 To 0)
    out(0) = ""
    GetCellGroupProperties = out
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging.GetCellGroupProperties"
End Function

' Read-only grammar validator for ONE rule (epic 12-2). Returns "" when the rule is valid (or empty,
' which the caller treats as a delete), otherwise a short English reason (fault/log channel). Purely
' SYNTACTIC: it enforces the structure of "selector=prop|prop" but does NOT check DGNLib membership (a
' property may be authored later). It reuses the SAME module constants as EnsureRulesParsed
' (SELECTOR_SEPARATOR / CELL_GROUP_MARKER / RULE_SEPARATOR / ARES_VAR_DELIMITER) so the validator and the
' parser cannot drift on what a rule looks like. It rejects exactly the malformations that fail silently
' in the parser today - above all a property token containing "=", "@" or ";" (the signature of using
' "|" instead of ";" between rules, the live bug this fixes). Standard error pattern -> fail-closed
' ("invalid rule"). Called from the options form on every commit (combo + raw reveal).
Public Function ValidateRuleSyntax(ByVal sRule As String) As String
    On Error GoTo ErrorHandler

    ValidateRuleSyntax = ""

    Dim s As String
    s = Trim(sRule)
    If Len(s) = 0 Then Exit Function

    Dim nEq As Long
    nEq = InStr(s, SELECTOR_SEPARATOR)
    If nEq = 0 Then
        ValidateRuleSyntax = "rule has no '=' (expected selector=prop|prop)"
        Exit Function
    End If

    Dim sSelector As String, sPropsRaw As String
    sSelector = Trim(Left(s, nEq - 1))
    sPropsRaw = Trim(Mid(s, nEq + 1))

    If Len(sSelector) = 0 Then
        ValidateRuleSyntax = "empty selector (expected level[:type] or @CellName before '=')"
        Exit Function
    End If
    If Len(sPropsRaw) = 0 Then
        ValidateRuleSyntax = "empty property list (expected prop|prop after '=')"
        Exit Function
    End If

    ' Selector: a cell rule (@CellName) needs a non-empty name with no stray "@" inside it; a level
    ' selector must not carry a stray "@" (that marker only leads a cell rule).
    If Left(sSelector, 1) = CELL_GROUP_MARKER Then
        Dim sCell As String
        sCell = Trim(Mid(sSelector, 2))
        If Len(sCell) = 0 Then
            ValidateRuleSyntax = "empty cell name after '@'"
            Exit Function
        End If
        If InStr(sCell, CELL_GROUP_MARKER) > 0 Then
            ValidateRuleSyntax = "stray '@' inside the cell name"
            Exit Function
        End If
    Else
        If InStr(sSelector, CELL_GROUP_MARKER) > 0 Then
            ValidateRuleSyntax = "stray '@' in a level selector (a cell rule must start with '@')"
            Exit Function
        End If
    End If

    ' Properties: each "|"-split token must be non-empty and free of "=", "@", ";". A token carrying "="
    ' or "@" almost always means the user typed "|" where a ";" (rule separator) was meant - name that.
    Dim vProps As Variant
    Dim i As Long, nValid As Long
    Dim sTok As String
    vProps = Split(sPropsRaw, ARESConstants.ARES_VAR_DELIMITER)
    nValid = 0
    For i = LBound(vProps) To UBound(vProps)
        sTok = Trim(vProps(i))
        If Len(sTok) > 0 Then
            If InStr(sTok, SELECTOR_SEPARATOR) > 0 Then
                ValidateRuleSyntax = "property '" & sTok & "' contains '=' - did you separate rules with '|' instead of ';'?"
                Exit Function
            End If
            If InStr(sTok, CELL_GROUP_MARKER) > 0 Then
                ValidateRuleSyntax = "property '" & sTok & "' contains '@' - did you separate rules with '|' instead of ';'?"
                Exit Function
            End If
            If InStr(sTok, RULE_SEPARATOR) > 0 Then
                ValidateRuleSyntax = "property '" & sTok & "' contains ';'"
                Exit Function
            End If
            nValid = nValid + 1
        End If
    Next i

    If nValid = 0 Then
        ValidateRuleSyntax = "no valid property"
        Exit Function
    End If

    ValidateRuleSyntax = ""
    Exit Function

ErrorHandler:
    ValidateRuleSyntax = "invalid rule"
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging.ValidateRuleSyntax"
End Function

'######################################################################################################################
'                                          PRIVATE HELPERS
'######################################################################################################################

' Does a single rule match the element? Level (case-insensitive) + type (only if the rule sets one).
Private Function RuleMatches(ByRef oRule As RuleInfo, ByVal oElement As element, ByVal sLevel As String) As Boolean
    On Error GoTo ErrorHandler

    RuleMatches = False
    If StrComp(oRule.Level, sLevel, vbTextCompare) <> 0 Then Exit Function
    If oRule.HasType Then
        If oElement.Type <> oRule.ElType Then Exit Function
    End If
    RuleMatches = True
    Exit Function

ErrorHandler:
    RuleMatches = False
End Function

' Parse ARES_Property_Rules into mRules once; cached until RefreshRules.
Private Sub EnsureRulesParsed()
    On Error GoTo ErrorHandler

    If mbParsed Then Exit Sub
    mbParsed = True
    mnRuleCount = 0

    Dim sRaw As String
    sRaw = GetRulesRaw()
    If Len(Trim(sRaw)) = 0 Then Exit Sub

    Dim vRules As Variant
    Dim k As Long, nEq As Long, nColon As Long
    Dim sSelector As String, sPropsRaw As String, sType As String
    Dim r As RuleInfo

    vRules = Split(sRaw, RULE_SEPARATOR)
    ReDim mRules(0 To UBound(vRules))

    For k = LBound(vRules) To UBound(vRules)
        If Len(Trim(vRules(k))) > 0 Then
            nEq = InStr(vRules(k), SELECTOR_SEPARATOR)
            If nEq > 0 Then
                sSelector = Trim(Left(vRules(k), nEq - 1))
                sPropsRaw = Trim(Mid(vRules(k), nEq + 1))

                If Left(sSelector, 1) = CELL_GROUP_MARKER Then
                    ' Cell-group rule: "@CellName=prop|prop". Properties attach to the OTHER members of
                    ' a graphic group containing a cell named CellName. A cell name has no type token,
                    ' so it is NOT split on ":". (r fields reset so a prior level rule cannot leak in.)
                    r.IsCellGroup = True
                    r.CellName = Trim(Mid(sSelector, 2))
                    r.Level = ""
                    r.HasType = False
                    If Len(r.CellName) > 0 And Len(sPropsRaw) > 0 Then
                        r.Props = SplitTrim(sPropsRaw, ARESConstants.ARES_VAR_DELIMITER)
                        mRules(mnRuleCount) = r
                        mnRuleCount = mnRuleCount + 1
                    End If
                Else
                    ' Level rule: selector = level[:type] (unchanged).
                    sType = ""
                    nColon = InStr(sSelector, TYPE_SEPARATOR)
                    If nColon > 0 Then
                        sType = Trim(Mid(sSelector, nColon + 1))
                        sSelector = Trim(Left(sSelector, nColon - 1))
                    End If

                    If Len(sSelector) > 0 And Len(sPropsRaw) > 0 Then
                        r.IsCellGroup = False
                        r.CellName = ""
                        r.Level = sSelector
                        r.HasType = (Len(sType) > 0)
                        If r.HasType Then r.ElType = MicroStationDefinition.StringToMsdElementType(sType)
                        r.Props = SplitTrim(sPropsRaw, ARESConstants.ARES_VAR_DELIMITER)
                        mRules(mnRuleCount) = r
                        mnRuleCount = mnRuleCount + 1
                    End If
                End If
            End If
        End If
    Next k
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging.EnsureRulesParsed"
    mnRuleCount = 0
End Sub

' Raw ARES_Property_Rules value ("" when unset). Lazily initialises ARESConfig like the other modules.
Private Function GetRulesRaw() As String
    On Error GoTo ErrorHandler

    GetRulesRaw = ""
    If ARESConfig Is Nothing Then Exit Function
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    If ARESConfig.ARES_PROPERTY_RULES Is Nothing Then Exit Function
    GetRulesRaw = ARESConfig.ARES_PROPERTY_RULES.Value
    Exit Function

ErrorHandler:
    GetRulesRaw = ""
End Function

' Split a string and trim each entry, dropping empties. Returns a 0-based array (a single "" when none).
Private Function SplitTrim(ByVal s As String, ByVal Delim As String) As String()
    Dim vParts As Variant, i As Long, n As Long
    Dim out() As String

    vParts = Split(s, Delim)
    ReDim out(0 To UBound(vParts))
    n = 0
    For i = LBound(vParts) To UBound(vParts)
        If Len(Trim(vParts(i))) > 0 Then
            out(n) = Trim(vParts(i))
            n = n + 1
        End If
    Next i

    If n = 0 Then
        ReDim out(0 To 0)
        out(0) = ""
    Else
        ReDim Preserve out(0 To n - 1)
    End If
    SplitTrim = out
End Function

' Safe "array has at least one element" check (mirrors ElementChangeHandler / PropertyPropagation).
' UBound returns -1 for an empty array and raises for an uninitialised one.
Private Function HasElements(ByRef arr() As element) As Boolean
    On Error Resume Next
    HasElements = False
    If UBound(arr) <> -1 Then HasElements = True
    On Error GoTo 0
End Function
