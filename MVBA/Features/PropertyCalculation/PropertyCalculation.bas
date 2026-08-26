' Module: PropertyCalculation
' Description: The VALUE-CALCULATION engine for ARES custom properties. Computes each custom property's
'              VALUE from a per-property calc rule (ARES_Calc_Rules) and writes it - but ONLY where that
'              property is ALREADY ATTACHED (by PropertyTagging). Never attaches/detaches directly.
'              Full grammar, engine passes and coupling doctrine: see _bmad/docs/calc-rules-grammar.md.
'              This module is the Core - rule cache, grammar/parsing, canonicalisation, Public API,
'              orchestration AND reporting (no dedicated Reporting module here, unlike PropertyRendering -
'              see the split plan §1/§4 for why: only 6 one-shot flags, and the sole read-accessor
'              (MultipleGeometriesReported) is externally pinned to Core anyway). Source evaluation lives
'              in PropertyCalculation_SourceEval; trigger-cell/trigger-level push mechanics live in
'              PropertyCalculation_TriggerPush.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (global ARESConfig), RuleGrammar, CustomPropertyHandler,
'               PropertyTagging, PropertyRendering, PropertyActuator, LangManager,
'               ErrorHandlerClass (global ErrorHandler), CallStackClass (global CallStack),
'               PropertyCalculation_SourceEval, PropertyCalculation_TriggerPush
'
' NOTE: ApplyValueToSibling calls PropertyActuator.ProcessElement on a sibling it just pushed a value to -
' a two-way coupling (PropertyActuator depends back on this module for IsTriggerCell/GetCalcRuleForProperty).
' Not a layering violation: ApplyValueToSibling is the only code that knows which sibling just changed.
' This coupling is orthogonal to the split and MUST stay intact/Public - do not move IsTriggerCell/
' IsTriggerLevel/GetCalcRuleForProperty out of this module.

Option Explicit

Private Const RULE_SEPARATOR As String = ";"
Private Const COND_SEPARATOR As String = "&"
Private Const SELECTOR_SEPARATOR As String = "="
Private Const BRK_OPEN As String = "["
Private Const BRK_CLOSE As String = "]"
Private Const PROP_KEYWORD As String = "PROP"
' Upper bound accepted for Coord[n]/Length[n]/GroupLength[n] decimal counts (syntactic; runtime formatting
' clamps to a sane max). Also within Length.GetLength's Byte RND range (255 is its reserved error sentinel).
Private Const SOURCE_MAX_DECIMALS As Long = 254

' Source vocabulary of a calc rule's right-hand side. csCell*/csLvl* are GROUP sources (self-included
' member scan by cell name / level name); the rest are SELF sources except csGroupLength (GROUP, by
' TYPE) and csGroupColor (GROUP, self-EXCLUDED - see EvaluateGroupColor).
Public Enum CalcSource
    csCellText
    csCellCoord
    csCellId
    csValue
    csCoord
    csId
    csLvl
    csCellLvl
    csColor
    csCellColor
    csStyle
    csCellStyle
    csWeight
    csCellWeight
    csLvlColor
    csLvlStyle
    csLvlWeight
    csGroupColor
    csLength
    csGroupLength
End Enum

' One parsed calc rule: Prop[TargetProp] [& conditions]* = Source. Conditions() (RuleGrammar.RuleCondition)
' is bounded by nCond. SourceArg holds the pattern (CellText/CellCoord/CellId), the fixed text (Value), the
' decimals string (Coord[n]) - empty for Id and bare Coord. Public: Module B (SourceEval) and Module C
' (TriggerPush) receive/return it across the module boundary.
Public Type CalcRuleInfo
    TargetProp As String
    Conditions() As RuleGrammar.RuleCondition
    nCond As Long
    SourceKind As CalcSource
    SourceArg As String
End Type

Private mCalcRules() As CalcRuleInfo
Private mnCalcCount As Long
Private mbCalcParsed As Boolean

' One-shot guards so each calculation status surfaces only once per processed element; reset at the
' start of ProcessElement. The four "Multiple" guards are DELIBERATELY SEPARATE, not one shared flag: a
' group can hold competing trigger cells, trigger levels, GroupColor candidates and geometries all at
' once, and a shared flag would silently swallow whichever fired second.
Private mbRejectedShown As Boolean
Private mbNoTargetShown As Boolean
Private mbMultiShown As Boolean
Private mbMultiGeoShown As Boolean
Private mbMultiLvlShown As Boolean
Private mbMultiColorShown As Boolean

'######################################################################################################################
'                                          PUBLIC SURFACE
'######################################################################################################################

' True once CalculationMultipleGeometries has been surfaced for the element being processed. Public
' (read-only) as a test seam - VBA cannot call a Private proc from UnitTesting.
Public Function MultipleGeometriesReported() As Boolean
    MultipleGeometriesReported = mbMultiGeoShown
End Function

' Master switch. Lazily initialises ARESConfig like the other feature modules.
Public Function IsEnabled() As Boolean
    On Error GoTo ErrorHandler

    IsEnabled = False
    If ARESConfig Is Nothing Then Exit Function
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    If ARESConfig.ARES_PROPERTY_CALC Is Nothing Then Exit Function
    IsEnabled = CBool(ARESConfig.ARES_PROPERTY_CALC.Value)
    Exit Function

ErrorHandler:
    IsEnabled = False
End Function

' Force a re-parse of ARES_Calc_Rules on the next resolve/apply (call after editing the variable).
Public Sub RefreshCalcRules()
    mbCalcParsed = False
End Sub

' Read-only validate-AND-normalise for ONE calc rule (the options editor's commit seam). Returns "" +
' canonical form on valid, "" + "" on empty (delete), or a short reason on invalid. Uses the SAME
' ParseCalcRule as the runtime, so it accepts exactly what the parser accepts.
Public Function ValidateAndNormalizeCalcRule(ByVal sRule As String, ByRef sCanonical As String) As String
    On Error GoTo ErrorHandler

    ValidateAndNormalizeCalcRule = ""
    sCanonical = ""

    Dim s As String
    s = Trim(sRule)
    If Len(s) = 0 Then Exit Function

    Dim r As CalcRuleInfo
    Dim sReason As String
    sReason = ParseCalcRule(s, r)
    If Len(sReason) > 0 Then
        ValidateAndNormalizeCalcRule = sReason
        Exit Function
    End If

    sCanonical = CalcRuleToCanonical(r)
    Exit Function

ErrorHandler:
    ValidateAndNormalizeCalcRule = "invalid rule"
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation.ValidateAndNormalizeCalcRule"
End Function

' Read-only contradiction (dead-rule) detector on a SYNTACTICALLY VALID calc rule. Returns True (with the
' two conflicting CONDITION segments, canonical text) when the conditions can never be satisfied together.
' The Prop target is ignored - only the conditions are reasoned over (delegated to
' RuleGrammar.ConditionsHaveContradiction, the same coverage as the tag detector). Used by the 14-3 preview.
Public Function CalcRuleHasNoEffect(ByVal sRule As String, ByRef segments() As String) As Boolean
    On Error GoTo ErrorHandler

    CalcRuleHasNoEffect = False
    ReDim segments(0 To 0)
    segments(0) = ""

    Dim r As CalcRuleInfo
    Dim sReason As String
    sReason = ParseCalcRule(sRule, r)
    If Len(sReason) > 0 Then Exit Function       ' only meaningful on a valid rule

    CalcRuleHasNoEffect = RuleGrammar.ConditionsHaveContradiction(r.Conditions, r.nCond, segments)
    Exit Function

ErrorHandler:
    ' Silent fail-closed (no log): a fault here only withholds an advisory verdict.
    CalcRuleHasNoEffect = False
    ReDim segments(0 To 0)
    segments(0) = ""
End Function

' A trigger cell is a CELL, in a REAL graphic group, whose name matches a pushable Cell* source's pattern.
' Drives the trigger-cell push pass; CellId is excluded (an ID never changes). An ungrouped matching cell
' is NOT a trigger - its own value is handled by the bearing pass instead.
Public Function IsTriggerCell(ByVal oEl As element) As Boolean
    On Error GoTo ErrorHandler

    IsTriggerCell = False
    If oEl Is Nothing Then Exit Function
    If Not oEl.IsCellElement Then Exit Function
    If oEl.GraphicGroup = ARES_DEFAULT_GRAPHIC_GROUP_ID Then Exit Function

    EnsureCalcRulesParsed
    If mnCalcCount = 0 Then Exit Function

    IsTriggerCell = PropertyCalculation_TriggerPush.AnyPushableSourcePatternMatches(oEl.AsCellElement.Name)
    Exit Function

ErrorHandler:
    IsTriggerCell = False
End Function

' Trigger test, mirrors IsTriggerCell but for a LEVEL match instead of a CELL-name match: oEl is a trigger
' when its OWN Level's name matches a pushable Lvl* source's pattern of at least one calc rule. NO element-
' type restriction (unlike IsTriggerCell's IsCellElement gate) - a Line/Arc on a matching level is a trigger
' just as much as a cell would be.
Public Function IsTriggerLevel(ByVal oEl As element) As Boolean
    On Error GoTo ErrorHandler

    IsTriggerLevel = False
    If oEl Is Nothing Then Exit Function
    If Not oEl.IsGraphical Then Exit Function
    If oEl.Level Is Nothing Then Exit Function
    If oEl.GraphicGroup = ARES_DEFAULT_GRAPHIC_GROUP_ID Then Exit Function

    EnsureCalcRulesParsed
    If mnCalcCount = 0 Then Exit Function

    IsTriggerLevel = PropertyCalculation_TriggerPush.AnyPushableLvlSourcePatternMatches(oEl.Level.Name)
    Exit Function

ErrorHandler:
    IsTriggerLevel = False
End Function

' Depth-0 hook: (1) BEARING pass - fill/recompute each calc-target property oEl carries; (2) TRIGGER-CELL
' pass - if oEl is a trigger cell, push its attributes to the OTHER members MicroStation didn't re-queue;
' (3) TRIGGER-LEVEL pass - same, for a level-bearing trigger of any type. Every write routes through
' ApplyValueToSibling (frontier + compare-guard + transition guard).
Public Sub ProcessElement(ByVal oEl As element)
    On Error GoTo ErrorHandler
    Dim bStackPushed As Boolean

    mbRejectedShown = False
    mbNoTargetShown = False
    mbMultiShown = False
    mbMultiGeoShown = False
    mbMultiLvlShown = False
    mbMultiColorShown = False

    If oEl Is Nothing Then Exit Sub

    EnsureCalcRulesParsed
    If mnCalcCount = 0 Then Exit Sub

    CallStack.Push "PropertyCalculation.ProcessElement", oEl
    bStackPushed = True

    BearingPass oEl

    If IsTriggerCell(oEl) Then PropertyCalculation_TriggerPush.PushCellDerivedValuesToMembers oEl
    If IsTriggerLevel(oEl) Then PropertyCalculation_TriggerPush.PushLvlDerivedValuesToMembers oEl
    CallStack.Pop
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation.ProcessElement"
    If bStackPushed Then CallStack.Pop
End Sub

' True when at least one calc rule uses GroupColor. Gates ElementChangeHandler's Branch 2 (mirrors
' HasGroupLengthRules): GroupColor has no name/level "trigger" pattern, so it needs every sibling
' re-evaluated inline when the linked element changes - Branch 2 already does that.
Public Function HasGroupColorRules() As Boolean
    On Error GoTo ErrorHandler

    HasGroupColorRules = False
    If Not IsEnabled Then Exit Function

    EnsureCalcRulesParsed
    Dim i As Long
    For i = 0 To mnCalcCount - 1
        If mCalcRules(i).SourceKind = csGroupColor Then
            HasGroupColorRules = True
            Exit Function
        End If
    Next i
    Exit Function

ErrorHandler:
    HasGroupColorRules = False
End Function

' True when at least one parsed calc rule uses GroupLength. Consulted by ElementChangeHandler's geometric
' Branch 2 gate: a GroupLength rule needs every OTHER group member re-queued when the linked geometry itself
' changes, same as Auto Lengths' own ARES_Update_Lengths gate - but must work even with Auto Lengths OFF.
' Public so ElementChangeHandler can short-circuit without duplicating the calc-rules cache.
Public Function HasGroupLengthRules() As Boolean
    On Error GoTo ErrorHandler

    HasGroupLengthRules = False
    If Not IsEnabled Then Exit Function

    EnsureCalcRulesParsed
    Dim i As Long
    For i = 0 To mnCalcCount - 1
        If mCalcRules(i).SourceKind = csGroupLength Then
            HasGroupLengthRules = True
            Exit Function
        End If
    Next i
    Exit Function

ErrorHandler:
    HasGroupLengthRules = False
End Function

' Number of parsed calc rules currently cached. Public accessor so Module C (TriggerPush) never reads
' mCalcRules()/mnCalcCount directly - added by the split (not in the original file).
Public Function RuleCount() As Long
    RuleCount = mnCalcCount
End Function

' One parsed calc rule by index, returned BY VALUE. Public accessor so Module C (TriggerPush) never reads
' the raw array directly - added by the split (not in the original file). Cheap: CalcRuleInfo is small.
Public Function GetRule(ByVal idx As Long) As CalcRuleInfo
    GetRule = mCalcRules(idx)
End Function

'######################################################################################################################
'                                          CALC-RULE GRAMMAR (single source of truth)
'######################################################################################################################

' Parse ARES_Calc_Rules into mCalcRules once; cached until RefreshCalcRules. Splits the raw value on the
' depth-0 ";" (a ";" is only ever a rule separator - forbidden inside [...]), then parses each rule via
' ParseCalcRule; a rule that does not fit the grammar is SKIPPED fail-closed (not counted) and logs nothing
' (a stored bad rule is not a fault).
Private Sub EnsureCalcRulesParsed()
    On Error GoTo ErrorHandler

    If mbCalcParsed Then Exit Sub
    mbCalcParsed = True
    mnCalcCount = 0

    Dim sRaw As String
    sRaw = GetCalcRulesRaw()
    If Len(Trim(sRaw)) = 0 Then Exit Sub

    Dim vRules() As String
    vRules = RuleGrammar.SplitTopLevel(sRaw, RULE_SEPARATOR)
    ReDim mCalcRules(0 To UBound(vRules))

    Dim k As Long
    Dim r As CalcRuleInfo
    Dim sReason As String
    For k = LBound(vRules) To UBound(vRules)
        If Len(Trim(vRules(k))) > 0 Then
            sReason = ParseCalcRule(vRules(k), r)
            If Len(sReason) = 0 Then
                mCalcRules(mnCalcCount) = r
                mnCalcCount = mnCalcCount + 1
            End If
        End If
    Next k
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation.EnsureCalcRulesParsed"
    mnCalcCount = 0
End Sub

' Parse ONE calc rule "Prop[name][&condition]*=Source". Returns "" on success (fills r) or a targeted
' English reason. "=", "&" are STRUCTURAL only at bracket depth 0; inside [...] they are literal. The LEFT
' side's segment[0] is Prop[name] (parsed here); segments[1..] are conditions (RuleGrammar.ParseCondition).
' The RIGHT side is a single arity-checked Source.
Private Function ParseCalcRule(ByVal sInput As String, ByRef r As CalcRuleInfo) As String
    On Error GoTo ErrorHandler

    ParseCalcRule = ""

    ' Reset the target so a previous rule cannot leak in.
    r.TargetProp = ""
    r.nCond = 0
    Erase r.Conditions
    r.SourceKind = csCellText
    r.SourceArg = ""

    Dim s As String
    s = Trim(sInput)
    If Len(s) = 0 Then
        ParseCalcRule = "empty rule"
        Exit Function
    End If

    ' Bracket-balance pre-check across the whole rule (catches every malformed [...]).
    Dim depth As Long, i As Long, c As String
    depth = 0
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c = BRK_OPEN Then
            depth = depth + 1
        ElseIf c = BRK_CLOSE Then
            depth = depth - 1
            If depth < 0 Then
                ParseCalcRule = "malformed [...] (unbalanced ']')"
                Exit Function
            End If
        End If
    Next i
    If depth <> 0 Then
        ParseCalcRule = "malformed [...] (unbalanced '[')"
        Exit Function
    End If

    ' First depth-0 "=".
    Dim eqPos As Long
    eqPos = RuleGrammar.FindTopLevelChar(s, SELECTOR_SEPARATOR)
    If eqPos = 0 Then
        ParseCalcRule = "rule has no '=' (expected Prop[name][&conditions]=Source)"
        Exit Function
    End If

    Dim leftSide As String, rightSide As String
    leftSide = Trim(Left(s, eqPos - 1))
    rightSide = Trim(Mid(s, eqPos + 1))
    If Len(leftSide) = 0 Then
        ParseCalcRule = "empty target side (before '=')"
        Exit Function
    End If
    If Len(rightSide) = 0 Then
        ParseCalcRule = "empty source side (after '=')"
        Exit Function
    End If

    ' Split the LEFT side on the depth-0 "&": segment[0] = Prop[name], segments[1..] = conditions.
    Dim segs() As String
    segs = RuleGrammar.SplitTopLevel(leftSide, COND_SEPARATOR)

    Dim sTargetReason As String
    sTargetReason = ParsePropTarget(segs(LBound(segs)), r)
    If Len(sTargetReason) > 0 Then
        ParseCalcRule = sTargetReason
        Exit Function
    End If

    ' Conditions (segments 1..): each parsed via the shared RuleGrammar condition grammar.
    If UBound(segs) - LBound(segs) >= 1 Then
        ReDim r.Conditions(0 To UBound(segs) - LBound(segs) - 1)
        Dim cnd As RuleGrammar.RuleCondition
        Dim seg As String
        Dim nc As Long
        nc = 0
        For i = LBound(segs) + 1 To UBound(segs)
            seg = Trim(segs(i))
            If Len(seg) = 0 Then
                ParseCalcRule = "empty condition (a '&' with nothing beside it)"
                Exit Function
            End If
            Dim sCondReason As String
            sCondReason = RuleGrammar.ParseCondition(seg, cnd)
            If Len(sCondReason) > 0 Then
                ParseCalcRule = sCondReason
                Exit Function
            End If
            r.Conditions(nc) = cnd
            nc = nc + 1
        Next i
        r.nCond = nc
    End If

    ' RIGHT side: a single arity-checked Source.
    Dim sSourceReason As String
    sSourceReason = ParseSource(rightSide, r)
    If Len(sSourceReason) > 0 Then
        ParseCalcRule = sSourceReason
        Exit Function
    End If
    Exit Function

ErrorHandler:
    ParseCalcRule = "invalid rule"
End Function

' Parse the target segment "Prop[name]" into r.TargetProp. Returns "" on success or a reason. The name is
' non-empty, wildcard-free, and free of the structural chars ";"/"[" (a "]" cannot occur - the first "]"
' closes the bracket). The name is kept VERBATIM (trimmed) - no DGNLib membership check (frontier guards it).
Private Function ParsePropTarget(ByVal sSeg As String, ByRef r As CalcRuleInfo) As String
    On Error GoTo ErrorHandler

    ParsePropTarget = ""

    Dim seg As String
    seg = Trim(sSeg)

    Dim nOpen As Long, nClose As Long
    nOpen = InStr(seg, BRK_OPEN)
    If nOpen = 0 Then
        ParsePropTarget = "left side must start with Prop[name]"
        Exit Function
    End If
    nClose = InStr(seg, BRK_CLOSE)
    If nClose <= nOpen Then
        ParsePropTarget = "malformed Prop[...]"
        Exit Function
    End If
    If nClose <> Len(seg) Then
        ParsePropTarget = "unexpected text after ']' in Prop[...]"
        Exit Function
    End If

    Dim kw As String
    kw = Trim(Left(seg, nOpen - 1))
    If UCase(kw) <> PROP_KEYWORD Then
        ParsePropTarget = "left side must start with Prop[name]"
        Exit Function
    End If

    Dim nm As String
    nm = Trim(Mid(seg, nOpen + 1, nClose - nOpen - 1))
    If Len(nm) = 0 Then
        ParsePropTarget = "empty property name in Prop[...]"
        Exit Function
    End If
    If InStr(nm, "*") > 0 Then
        ParsePropTarget = "wildcards not allowed in Prop[...]"
        Exit Function
    End If
    If InStr(nm, "?") > 0 Then
        ParsePropTarget = "wildcards not allowed in Prop[...]"
        Exit Function
    End If
    If InStr(nm, RULE_SEPARATOR) > 0 Then
        ParsePropTarget = "';' not allowed in Prop[...]"
        Exit Function
    End If
    If InStr(nm, BRK_OPEN) > 0 Then
        ParsePropTarget = "'[' not allowed in Prop[...]"
        Exit Function
    End If

    r.TargetProp = nm
    Exit Function

ErrorHandler:
    ParsePropTarget = "invalid Prop[...] target"
End Function

' Parse the RIGHT side into a Source (keyword + optional [arg]), arity-checked, unknown rejected. Fills
' r.SourceKind / r.SourceArg. Returns "" on success or a targeted reason.
Private Function ParseSource(ByVal sRight As String, ByRef r As CalcRuleInfo) As String
    On Error GoTo ErrorHandler

    ParseSource = ""

    Dim src As String
    src = Trim(sRight)

    Dim kw As String, arg As String
    Dim bHasArg As Boolean
    Dim nOpen As Long, nClose As Long
    nOpen = InStr(src, BRK_OPEN)
    If nOpen = 0 Then
        kw = src
        arg = ""
        bHasArg = False
    Else
        nClose = InStr(src, BRK_CLOSE)
        If nClose <= nOpen Then
            ParseSource = "malformed source [...]"
            Exit Function
        End If
        If nClose <> Len(src) Then
            ParseSource = "unexpected text after ']' in source"
            Exit Function
        End If
        kw = Trim(Left(src, nOpen - 1))
        arg = Mid(src, nOpen + 1, nClose - nOpen - 1)
        bHasArg = True
        ' Structural chars forbidden inside a source arg (mirrors the condition grammar).
        If InStr(arg, RULE_SEPARATOR) > 0 Then
            ParseSource = "';' not allowed inside [...]"
            Exit Function
        End If
        If InStr(arg, BRK_OPEN) > 0 Then
            ParseSource = "'[' not allowed inside [...]"
            Exit Function
        End If
    End If

    Dim pat As String
    Dim sPatReason As String

    Select Case UCase(kw)
        Case "CELLTEXT"
            If Not RequirePatternArg("CellText", bHasArg, arg, pat, sPatReason) Then
                ParseSource = sPatReason
                Exit Function
            End If
            r.SourceKind = csCellText
            r.SourceArg = pat
        Case "CELLCOORD"
            If Not RequirePatternArg("CellCoord", bHasArg, arg, pat, sPatReason) Then
                ParseSource = sPatReason
                Exit Function
            End If
            r.SourceKind = csCellCoord
            r.SourceArg = pat
        Case "CELLID"
            If Not RequirePatternArg("CellId", bHasArg, arg, pat, sPatReason) Then
                ParseSource = sPatReason
                Exit Function
            End If
            r.SourceKind = csCellId
            r.SourceArg = pat
        Case "CELLLVL"
            If Not RequirePatternArg("CellLvl", bHasArg, arg, pat, sPatReason) Then
                ParseSource = sPatReason
                Exit Function
            End If
            r.SourceKind = csCellLvl
            r.SourceArg = pat
        Case "CELLCOLOR"
            If Not RequirePatternArg("CellColor", bHasArg, arg, pat, sPatReason) Then
                ParseSource = sPatReason
                Exit Function
            End If
            r.SourceKind = csCellColor
            r.SourceArg = pat
        Case "CELLSTYLE"
            If Not RequirePatternArg("CellStyle", bHasArg, arg, pat, sPatReason) Then
                ParseSource = sPatReason
                Exit Function
            End If
            r.SourceKind = csCellStyle
            r.SourceArg = pat
        Case "CELLWEIGHT"
            If Not RequirePatternArg("CellWeight", bHasArg, arg, pat, sPatReason) Then
                ParseSource = sPatReason
                Exit Function
            End If
            r.SourceKind = csCellWeight
            r.SourceArg = pat
        Case "LVLCOLOR"
            If Not RequirePatternArg("LvlColor", bHasArg, arg, pat, sPatReason) Then
                ParseSource = sPatReason
                Exit Function
            End If
            r.SourceKind = csLvlColor
            r.SourceArg = pat
        Case "LVLSTYLE"
            If Not RequirePatternArg("LvlStyle", bHasArg, arg, pat, sPatReason) Then
                ParseSource = sPatReason
                Exit Function
            End If
            r.SourceKind = csLvlStyle
            r.SourceArg = pat
        Case "LVLWEIGHT"
            If Not RequirePatternArg("LvlWeight", bHasArg, arg, pat, sPatReason) Then
                ParseSource = sPatReason
                Exit Function
            End If
            r.SourceKind = csLvlWeight
            r.SourceArg = pat
        Case "VALUE"
            If Not bHasArg Then
                ParseSource = "Value needs a [text]"
                Exit Function
            End If
            If Len(arg) = 0 Then
                ParseSource = "empty Value[...]"
                Exit Function
            End If
            r.SourceKind = csValue
            r.SourceArg = arg                    ' verbatim: a fixed value keeps its content exactly
        Case "COORD"
            r.SourceKind = csCoord
            r.SourceArg = ""
            If bHasArg Then
                Dim sN As String
                sN = Trim(arg)
                If Not IsNonNegIntInRange(sN, 0, SOURCE_MAX_DECIMALS) Then
                    ParseSource = "Coord[n] needs an integer decimal count"
                    Exit Function
                End If
                r.SourceArg = sN
            End If
        Case "ID"
            If bHasArg Then
                ParseSource = "Id takes no argument"
                Exit Function
            End If
            r.SourceKind = csId
            r.SourceArg = ""
        Case "LVL"
            If bHasArg Then
                ParseSource = "Lvl takes no argument"
                Exit Function
            End If
            r.SourceKind = csLvl
            r.SourceArg = ""
        Case "COLOR"
            If bHasArg Then
                ParseSource = "Color takes no argument"
                Exit Function
            End If
            r.SourceKind = csColor
            r.SourceArg = ""
        Case "STYLE"
            If bHasArg Then
                ParseSource = "Style takes no argument"
                Exit Function
            End If
            r.SourceKind = csStyle
            r.SourceArg = ""
        Case "WEIGHT"
            If bHasArg Then
                ParseSource = "Weight takes no argument"
                Exit Function
            End If
            r.SourceKind = csWeight
            r.SourceArg = ""
        Case "GROUPCOLOR"
            If bHasArg Then
                ParseSource = "GroupColor takes no argument"
                Exit Function
            End If
            r.SourceKind = csGroupColor
            r.SourceArg = ""
        Case "LENGTH"
            r.SourceKind = csLength
            r.SourceArg = ""
            If bHasArg Then
                Dim sNLen As String
                sNLen = Trim(arg)
                If Not IsNonNegIntInRange(sNLen, 0, SOURCE_MAX_DECIMALS) Then
                    ParseSource = "Length[n] needs an integer decimal count"
                    Exit Function
                End If
                r.SourceArg = sNLen
            End If
        Case "GROUPLENGTH"
            r.SourceKind = csGroupLength
            r.SourceArg = ""
            If bHasArg Then
                Dim sNGrp As String
                sNGrp = Trim(arg)
                If Not IsNonNegIntInRange(sNGrp, 0, SOURCE_MAX_DECIMALS) Then
                    ParseSource = "GroupLength[n] needs an integer decimal count"
                    Exit Function
                End If
                r.SourceArg = sNGrp
            End If
        Case Else
            If Len(kw) = 0 Then
                ParseSource = "empty source (expected CellText/CellCoord/CellId/CellLvl/CellColor/CellStyle/CellWeight/LvlColor/LvlStyle/LvlWeight/GroupColor/Value/Coord/Id/Lvl/Color/Style/Weight/Length/GroupLength)"
            Else
                ParseSource = "unknown source '" & kw & "' (expected CellText/CellCoord/CellId/CellLvl/CellColor/CellStyle/CellWeight/LvlColor/LvlStyle/LvlWeight/GroupColor/Value/Coord/Id/Lvl/Color/Style/Weight/Length/GroupLength)"
            End If
            Exit Function
    End Select
    Exit Function

ErrorHandler:
    ParseSource = "invalid source"
End Function

' Shared arity check for every GROUP source that takes a mandatory non-empty [pattern] (CellText/CellCoord/
' CellId/CellLvl/CellColor/CellStyle/CellWeight/LvlColor/LvlStyle/LvlWeight). Returns False + sReason on a
' missing/empty pattern, True + outPat (trimmed) on success.
Private Function RequirePatternArg(ByVal sKeyword As String, ByVal bHasArg As Boolean, ByVal arg As String, ByRef outPat As String, ByRef sReason As String) As Boolean
    RequirePatternArg = False
    outPat = Trim(arg)
    If Not bHasArg Then
        sReason = sKeyword & " needs a [pattern]"
        Exit Function
    End If
    If Len(outPat) = 0 Then
        sReason = "empty " & sKeyword & "[...] pattern"
        Exit Function
    End If
    RequirePatternArg = True
End Function

' True when s is a non-negative integer (all digits, at least one) whose value is in [lo, hi].
Private Function IsNonNegIntInRange(ByVal s As String, ByVal lo As Long, ByVal hi As Long) As Boolean
    On Error GoTo ErrorHandler

    IsNonNegIntInRange = False
    If Len(s) = 0 Then Exit Function

    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch < "0" Then Exit Function
        If ch > "9" Then Exit Function
    Next i

    Dim v As Long
    v = CLng(s)
    If v >= lo Then
        If v <= hi Then IsNonNegIntInRange = True
    End If
    Exit Function

ErrorHandler:
    IsNonNegIntInRange = False
End Function

'######################################################################################################################
'                                          CANONICALISATION
'######################################################################################################################

' Build the COMPACT canonical form of a parsed calc rule: Prop[name] [&cond]* = Source, no spaces around
' "&"/"=", canonical Prop/source keyword casing, condition text via RuleGrammar.ConditionToCanonical,
' names/args verbatim.
Private Function CalcRuleToCanonical(ByRef r As CalcRuleInfo) As String
    Dim sOut As String
    Dim i As Long

    sOut = "Prop" & BRK_OPEN & r.TargetProp & BRK_CLOSE

    For i = 0 To r.nCond - 1
        sOut = sOut & COND_SEPARATOR & RuleGrammar.ConditionToCanonical(r.Conditions(i))
    Next i

    sOut = sOut & SELECTOR_SEPARATOR & SourceToCanonical(r)
    CalcRuleToCanonical = sOut
End Function

' Canonical text of a rule's Source (keyword canonical, arg verbatim). Bracketed-arg kinds (Cell* and the
' optional-decimals Coord/Length/GroupLength) share one bracket-wrap helper.
Private Function SourceToCanonical(ByRef r As CalcRuleInfo) As String
    Select Case r.SourceKind
        Case csCellText
            SourceToCanonical = "CellText" & BRK_OPEN & r.SourceArg & BRK_CLOSE
        Case csCellCoord
            SourceToCanonical = "CellCoord" & BRK_OPEN & r.SourceArg & BRK_CLOSE
        Case csCellId
            SourceToCanonical = "CellId" & BRK_OPEN & r.SourceArg & BRK_CLOSE
        Case csCellLvl
            SourceToCanonical = "CellLvl" & BRK_OPEN & r.SourceArg & BRK_CLOSE
        Case csCellColor
            SourceToCanonical = "CellColor" & BRK_OPEN & r.SourceArg & BRK_CLOSE
        Case csCellStyle
            SourceToCanonical = "CellStyle" & BRK_OPEN & r.SourceArg & BRK_CLOSE
        Case csCellWeight
            SourceToCanonical = "CellWeight" & BRK_OPEN & r.SourceArg & BRK_CLOSE
        Case csLvlColor
            SourceToCanonical = "LvlColor" & BRK_OPEN & r.SourceArg & BRK_CLOSE
        Case csLvlStyle
            SourceToCanonical = "LvlStyle" & BRK_OPEN & r.SourceArg & BRK_CLOSE
        Case csLvlWeight
            SourceToCanonical = "LvlWeight" & BRK_OPEN & r.SourceArg & BRK_CLOSE
        Case csValue
            SourceToCanonical = "Value" & BRK_OPEN & r.SourceArg & BRK_CLOSE
        Case csCoord
            SourceToCanonical = OptionalArgKeywordToCanonical("Coord", r.SourceArg)
        Case csId
            SourceToCanonical = "Id"
        Case csLvl
            SourceToCanonical = "Lvl"
        Case csColor
            SourceToCanonical = "Color"
        Case csStyle
            SourceToCanonical = "Style"
        Case csWeight
            SourceToCanonical = "Weight"
        Case csGroupColor
            SourceToCanonical = "GroupColor"
        Case csLength
            SourceToCanonical = OptionalArgKeywordToCanonical("Length", r.SourceArg)
        Case csGroupLength
            SourceToCanonical = OptionalArgKeywordToCanonical("GroupLength", r.SourceArg)
        Case Else
            SourceToCanonical = ""
    End Select
End Function

' sKeyword bare, or sKeyword[arg] when arg is non-empty - shared by Coord/Length/GroupLength (all take an
' OPTIONAL decimals arg, unlike the mandatory-pattern Cell* sources).
Private Function OptionalArgKeywordToCanonical(ByVal sKeyword As String, ByVal sArg As String) As String
    If Len(sArg) > 0 Then
        OptionalArgKeywordToCanonical = sKeyword & BRK_OPEN & sArg & BRK_CLOSE
    Else
        OptionalArgKeywordToCanonical = sKeyword
    End If
End Function

'######################################################################################################################
'                                          ENGINE - BEARING PASS
'######################################################################################################################

' Bearing pass: for each DISTINCT calc-target property P that oEl currently carries (frontier), resolve its
' value from the FIRST matching calc rule and write it (compare-guarded). A property with NO matching rule
' is left untouched (the engine only governs what a rule matches); a matching rule that yields "" (e.g. a
' CellText with no surviving cell) empties the value via the transition-guarded ApplyValueToSibling.
Private Sub BearingPass(ByVal oEl As element)
    On Error GoTo ErrorHandler

    Dim targets() As String
    Dim nT As Long
    targets = DistinctTargets(nT)
    If nT = 0 Then Exit Sub

    Dim i As Long
    Dim bHasRule As Boolean
    Dim sVal As String
    For i = 0 To nT - 1
        If CustomPropertyHandler.IsItemAttachedToElement(oEl, targets(i)) Then
            sVal = ResolvePropertyValue(targets(i), oEl, bHasRule)
            If bHasRule Then ApplyValueToSibling oEl, targets(i), sVal
        End If
    Next i
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation.BearingPass"
End Sub

' Resolves P's value from its first matching calc rule. bHasRule False = no rule governs P (caller leaves
' it alone); True with "" = the rule yields empty (caller clears/detaches). Read-only test/preview seam.
Public Function ResolvePropertyValue(ByVal P As String, ByVal oEl As element, ByRef bHasRule As Boolean) As String
    On Error GoTo ErrorHandler

    ResolvePropertyValue = ""
    bHasRule = False

    Dim idx As Long
    idx = FindCalcRuleForProperty(P, oEl)
    If idx < 0 Then Exit Function

    bHasRule = True
    ResolvePropertyValue = PropertyCalculation_SourceEval.EvaluateSource(mCalcRules(idx), oEl)
    Exit Function

ErrorHandler:
    ' Fail-closed: withhold a value AND signal "no rule" so the caller does not clear on a fault.
    ResolvePropertyValue = ""
    bHasRule = False
End Function

' Read-only "where does P's value come from on this element" seam - False when no rule governs P.
' PropertyRendering uses it at BIND time to detect the one static cycle v1 can catch.
Public Function GetCalcRuleForProperty(ByVal P As String, ByVal oEl As element, ByRef SourceKind As CalcSource, ByRef SourceArg As String, ByRef sCanonical As String) As Boolean
    On Error GoTo ErrorHandler

    GetCalcRuleForProperty = False
    SourceArg = ""
    sCanonical = ""

    EnsureCalcRulesParsed
    If mnCalcCount = 0 Then Exit Function

    Dim idx As Long
    idx = FindCalcRuleForProperty(P, oEl)
    If idx < 0 Then Exit Function

    SourceKind = mCalcRules(idx).SourceKind
    SourceArg = mCalcRules(idx).SourceArg
    sCanonical = SourceToCanonical(mCalcRules(idx))
    GetCalcRuleForProperty = True
    Exit Function

ErrorHandler:
    GetCalcRuleForProperty = False
End Function

' First-match resolution: index of the first calc rule whose TargetProp = P (case-insensitive) AND whose
' conditions match oEl; -1 if none. Order in the config = priority (specific rules first). Public: called
' from Module C (TriggerPush)'s Push*DerivedValuesToMembers (AC3 first-match guard).
Public Function FindCalcRuleForProperty(ByVal P As String, ByVal oEl As element) As Long
    On Error GoTo ErrorHandler

    FindCalcRuleForProperty = -1
    If oEl Is Nothing Then Exit Function

    ' Resolve the level once (guarded). A cell header is graphical but has no Level; Cell/Type conditions
    ' must still be evaluated, so we pass bHasLevel = False rather than exit.
    Dim sLevel As String
    Dim bHasLevel As Boolean
    sLevel = ""
    bHasLevel = False
    If oEl.IsGraphical Then
        If Not oEl.Level Is Nothing Then
            sLevel = oEl.Level.Name
            bHasLevel = True
        End If
    End If

    Dim i As Long
    For i = 0 To mnCalcCount - 1
        If StrComp(mCalcRules(i).TargetProp, P, vbTextCompare) = 0 Then
            If RuleMatchesConditions(mCalcRules(i), oEl, sLevel, bHasLevel) Then
                FindCalcRuleForProperty = i
                Exit Function
            End If
        End If
    Next i
    Exit Function

ErrorHandler:
    FindCalcRuleForProperty = -1
End Function

' AND over a rule's conditions with strict negation (each via RuleGrammar.ConditionMatches). No conditions
' -> matches everything. sLevel/bHasLevel are resolved once by the caller (guarded).
Private Function RuleMatchesConditions(ByRef r As CalcRuleInfo, ByVal oEl As element, ByVal sLevel As String, ByVal bHasLevel As Boolean) As Boolean
    On Error GoTo ErrorHandler

    RuleMatchesConditions = False
    Dim i As Long
    For i = 0 To r.nCond - 1
        If Not RuleGrammar.ConditionMatches(r.Conditions(i), oEl, sLevel, bHasLevel) Then Exit Function
    Next i
    RuleMatchesConditions = True
    Exit Function

ErrorHandler:
    RuleMatchesConditions = False
End Function

'######################################################################################################################
'                                          LOW-LEVEL HELPERS
'######################################################################################################################

' Distinct calc-target property names across the parsed rules (case-insensitive dedup). Returns a 0-based
' array (a single "" when none) and the count in nOut.
Private Function DistinctTargets(ByRef nOut As Long) As String()
    On Error GoTo ErrorHandler

    Dim out() As String
    Dim n As Long
    ReDim out(0 To 0)
    n = 0

    Dim i As Long, j As Long
    Dim bDup As Boolean
    For i = 0 To mnCalcCount - 1
        bDup = False
        For j = 0 To n - 1
            If StrComp(out(j), mCalcRules(i).TargetProp, vbTextCompare) = 0 Then
                bDup = True
                Exit For
            End If
        Next j
        If Not bDup Then
            If n > UBound(out) Then ReDim Preserve out(0 To n)
            out(n) = mCalcRules(i).TargetProp
            n = n + 1
        End If
    Next i

    If n = 0 Then
        ReDim out(0 To 0)
        out(0) = ""
    Else
        ReDim Preserve out(0 To n - 1)
    End If
    nOut = n
    DistinctTargets = out
    Exit Function

ErrorHandler:
    nOut = 0
    ReDim out(0 To 0)
    out(0) = ""
    DistinctTargets = out
End Function

' Raw ARES_Calc_Rules value ("" when unset). Lazily initialises ARESConfig like the other modules.
Private Function GetCalcRulesRaw() As String
    On Error GoTo ErrorHandler

    GetCalcRulesRaw = ""
    If ARESConfig Is Nothing Then Exit Function
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    If ARESConfig.ARES_CALC_RULES Is Nothing Then Exit Function
    GetCalcRulesRaw = ARESConfig.ARES_CALC_RULES.Value
    Exit Function

ErrorHandler:
    GetCalcRulesRaw = ""
End Function

'######################################################################################################################
'                        VALUE-WRITE MACHINERY (loop-safety BLOCKERs - do not weaken the guards)
'######################################################################################################################

' Frontier + compare-before-write on one sibling (loop-safety). Returns True when P is already attached
' (write attempted or not). Never attaches/detaches directly: an emptying (non-empty -> empty) TRANSITION
' either clears the value or, with ARES_Calc_Detach_Empty ON, delegates a detach to the tagger - the ONLY
' detach path, gated on the transition so a re-attaching rule cannot oscillate it. Public: called from
' Module C (TriggerPush)'s Push*DerivedValuesToMembers, and from BearingPass here.
Public Function ApplyValueToSibling(ByVal s As element, ByVal P As String, ByVal value As String) As Boolean
    On Error GoTo ErrorHandler

    ApplyValueToSibling = False
    If s Is Nothing Then Exit Function

    ' Frontier: write only where P is ALREADY attached (HasItems, not Null-inference - an attached-but-
    ' empty property also reads back Null). Not attached -> skip; attach stays the tagger's domain.
    If Not CustomPropertyHandler.IsItemAttachedToElement(s, P) Then Exit Function
    ApplyValueToSibling = True

    ' Split-coordinate items: detected by the TARGET ITEM'S SHAPE alone (a 2-field X/Y ItemType - see
    ' CustomPropertyHandler.GetXYSplitMembers), NOT by which SourceKind produced value. Any calc rule
    ' (Coord/CellCoord, but also Value[...] or any other source) that targets such an item takes this
    ' split-write path; ApplyXYValueToSibling rejects outright (writes nothing) if value is not exactly
    ' "part;part" once non-empty, rather than truncating silently. Every OTHER target (the 1-field shape
    ' every other property uses) takes the unchanged single-field path below unmodified - GetXYSplitMembers
    ' is False for a 1-member item by construction, so this is a pure no-op check for every property that
    ' isn't a split coordinate.
    Dim oItem As ItemType
    Dim sXMember As String, sYMember As String
    Set oItem = CustomPropertyHandler.GetItemTypeFromElement(s, P, ARESConstants.ARES_NAME_LIBRARY_TYPE)
    If CustomPropertyHandler.GetXYSplitMembers(oItem, sXMember, sYMember) Then
        ApplyXYValueToSibling s, P, sXMember, sYMember, value
        Exit Function
    End If

    ' Read the current value. Nested read-then-branch keeps CStr off a possible array (no short-circuit
    ' in VBA); an attached-but-empty property reads back Null -> sCurrent "".
    Dim vCurrent As Variant
    Dim sCurrent As String
    vCurrent = CustomPropertyHandler.GetPropertyValueFromElement(s, P, P)
    If IsNull(vCurrent) Then sCurrent = "" Else sCurrent = CStr(vCurrent)

    If Len(value) > 0 Then
        ' Non-empty value: set only when different (compare-guarded).
        If sCurrent <> value Then
            If CustomPropertyHandler.SetPropertyValueToElement(s, P, value, P) Then
                ' s was never queued by MicroStation (only the trigger was), so nothing else would refresh
                ' its rendered text or repaint its Color/Level from this fresh value - note/react inline.
                PropertyRendering.NoteDirtyGroup s
                PropertyActuator.ProcessElement s
            Else
                ReportRejected
            End If
        End If
        ' already equal -> no-op (loop-safety)
    Else
        ' Empty value: act ONLY on a real non-empty -> empty TRANSITION (an already-empty property is a
        ' no-op, so a rule that re-attaches P empty does not re-trigger a detach - this makes ON terminate).
        If Len(sCurrent) > 0 Then
            Dim bEmptied As Boolean
            bEmptied = True
            If IsDetachEmptyEnabled() Then
                ' Option ON: delegate the detach to the tagger (the only permitted detach path).
                PropertyTagging.DetachRuleProperty s, P
            Else
                ' Option OFF: clear the value; the property stays attached.
                bEmptied = CustomPropertyHandler.SetPropertyValueToElement(s, P, "", P)
                If Not bEmptied Then ReportRejected
            End If
            ' Gated on the clear/detach having been ACCEPTED - a rejected write warrants no repaint.
            If bEmptied Then
                PropertyRendering.NoteDirtyGroup s
                PropertyActuator.ProcessElement s
            End If
        End If
    End If
    Exit Function

ErrorHandler:
    ' A fault mid-write does not un-attach P; the return value (set True once past the frontier) is only
    ' used to detect "no member carried P", so leaving it as-is is correct.
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation.ApplyValueToSibling"
End Function

' Split-coordinate write path for ApplyValueToSibling (design: see plan-xy-split-coordinate-properties.md
' §3, Option C). value is the combined "X;Y" string a Coord/CellCoord source produced (FormatCoord's own
' field separator) - split as a plain STRING operation, never re-parsed as a number here (both target
' fields are String properties; no numeric round-trip is involved). Reads/compares/writes each field
' INDEPENDENTLY, each behind its own bNoFallback:=True call - mandatory, not optional: the fallback that
' resolves a mismatched access string by scanning ALL of an item's members (CustomPropertyHandler.
' GetFirstPropertyValue/SetFirstPropertyValue) would otherwise silently answer with/write the WRONG axis on
' a genuinely multi-property item (see the plan's §3.1). "Empty" for the transition guard means BOTH
' components empty - Coord/CellCoord's source (GetElementAnchorPoint) is always fully populated or not at
' all, never just one axis (see the plan's §4), so this is the only shape this path is ever reached with.
' A fault between the X write and the Y write (element becomes invalid mid-pass) is accepted as-is per the
' plan's §4/lead's decision - no extra IsElementValid re-check between the two writes; a later pass
' recomputes and converges, same as any other transient write failure this module already tolerates.
'
' The split is triggered by the TARGET ITEM'S SHAPE alone (GetXYSplitMembers), not by which SourceKind
' produced value - a Value[...] literal, not just Coord/CellCoord, reaching a 2-field X/Y item takes this
' path too (see the plan's §1: "any ItemType matching this shape, whatever its name"). A malformed value
' (not exactly 2 non-empty ";"-separated parts) is therefore possible here - e.g. Value[bonjour] (0 ";")
' or Value[A;B;C] (2 ";"). Per lead's decision, this is REJECTED outright (nothing written to either
' field) rather than silently truncated - "no value rather than a wrong value", consistent with
' IsProtectedTriggerCell/IsSelfSourceRatchet/GetElementAnchorPoint elsewhere in ARES. A truly empty value
' (value = "") is NOT malformed - it is the legitimate empty-transition case handled below.
Private Sub ApplyXYValueToSibling(ByVal s As element, ByVal P As String, ByVal sXMember As String, ByVal sYMember As String, ByVal value As String)
    On Error GoTo ErrorHandler

    Dim parts() As String
    Dim sNewX As String, sNewY As String
    If Len(value) > 0 Then
        parts = Split(value, ";")
        If (UBound(parts) - LBound(parts)) <> 1 Then
            ReportRejected                        ' not exactly 2 parts - malformed, write nothing
            Exit Sub
        End If
        sNewX = parts(LBound(parts))
        sNewY = parts(LBound(parts) + 1)
        If Len(sNewX) = 0 Or Len(sNewY) = 0 Then
            ReportRejected                        ' a part is empty ("X;" or ";Y") - malformed, write nothing
            Exit Sub
        End If
    End If

    Dim vCurX As Variant, vCurY As Variant
    Dim sCurX As String, sCurY As String
    vCurX = CustomPropertyHandler.GetPropertyValueFromElement(s, sXMember, P, ARESConstants.ARES_NAME_LIBRARY_TYPE, True)
    vCurY = CustomPropertyHandler.GetPropertyValueFromElement(s, sYMember, P, ARESConstants.ARES_NAME_LIBRARY_TYPE, True)
    If IsNull(vCurX) Then sCurX = "" Else sCurX = CStr(vCurX)
    If IsNull(vCurY) Then sCurY = "" Else sCurY = CStr(vCurY)

    If Len(sNewX) > 0 Or Len(sNewY) > 0 Then
        ' Non-empty: write whichever field differs (independent compare-guard per field - see the plan's
        ' §4; X/Y are never independently absent for THIS source, so in practice both differ together).
        Dim bChanged As Boolean
        bChanged = False
        If sCurX <> sNewX Then
            If CustomPropertyHandler.SetPropertyValueToElement(s, sXMember, sNewX, P, ARESConstants.ARES_NAME_LIBRARY_TYPE, True) Then
                bChanged = True
            Else
                ReportRejected
            End If
        End If
        If sCurY <> sNewY Then
            If CustomPropertyHandler.SetPropertyValueToElement(s, sYMember, sNewY, P, ARESConstants.ARES_NAME_LIBRARY_TYPE, True) Then
                bChanged = True
            Else
                ReportRejected
            End If
        End If
        If bChanged Then
            PropertyRendering.NoteDirtyGroup s
            PropertyActuator.ProcessElement s
        End If
    Else
        ' Empty value: act ONLY on a real non-empty -> empty TRANSITION, mirroring the single-field path.
        ' "Empty" here means BOTH components empty (see header comment).
        If Len(sCurX) > 0 Or Len(sCurY) > 0 Then
            Dim bEmptiedX As Boolean, bEmptiedY As Boolean
            bEmptiedX = True
            bEmptiedY = True
            If IsDetachEmptyEnabled() Then
                ' Option ON: delegate the detach to the tagger - item-scoped, not field-scoped, same as
                ' the single-field path (DetachRuleProperty already detaches the whole item today).
                PropertyTagging.DetachRuleProperty s, P
            Else
                If Len(sCurX) > 0 Then
                    bEmptiedX = CustomPropertyHandler.SetPropertyValueToElement(s, sXMember, "", P, ARESConstants.ARES_NAME_LIBRARY_TYPE, True)
                    If Not bEmptiedX Then ReportRejected
                End If
                If Len(sCurY) > 0 Then
                    bEmptiedY = CustomPropertyHandler.SetPropertyValueToElement(s, sYMember, "", P, ARESConstants.ARES_NAME_LIBRARY_TYPE, True)
                    If Not bEmptiedY Then ReportRejected
                End If
            End If
            ' Gated on BOTH clears having been accepted - a partial rejection still warrants a repaint
            ' (something did change), mirrored loosely on the single-field "gated on acceptance" rule.
            If bEmptiedX And bEmptiedY Then
                PropertyRendering.NoteDirtyGroup s
                PropertyActuator.ProcessElement s
            End If
        End If
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation.ApplyXYValueToSibling"
End Sub

' Option (ARES_Calc_Detach_Empty): when True, an emptied value is DETACHED (delegated to the tagger)
' instead of cleared. Mirrors IsEnabled - fail-closed False on any nil; lazy ARESConfig init.
Private Function IsDetachEmptyEnabled() As Boolean
    On Error GoTo ErrorHandler

    IsDetachEmptyEnabled = False
    If ARESConfig Is Nothing Then Exit Function
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    If ARESConfig.ARES_CALC_DETACH_EMPTY Is Nothing Then Exit Function
    IsDetachEmptyEnabled = CBool(ARESConfig.ARES_CALC_DETACH_EMPTY.Value)
    Exit Function

ErrorHandler:
    IsDetachEmptyEnabled = False
End Function

' Log the rejected write (English, Number 0) and surface CalculationValueRejected ONCE per processed element.
Private Sub ReportRejected()
    On Error Resume Next
    ErrorHandler.HandleError "Property calculation: target property rejected the value", 0, "", "PropertyCalculation.ApplyValueToSibling"
    If Not mbRejectedShown Then
        LangManager.ShowStatusT "CalculationValueRejected"
        mbRejectedShown = True
    End If
End Sub

' A trigger cell fired but no sibling carried its target property. Status-only, no log. Public: called
' from Module C (TriggerPush)'s Push*DerivedValuesToMembers.
Public Sub ReportNoTarget()
    On Error Resume Next
    If Not mbNoTargetShown Then
        LangManager.ShowStatusT "CalculationNoTarget"
        mbNoTargetShown = True
    End If
End Sub

' Two competing trigger cells for one target. Status-only, no log. Last-processed wins. Public: called
' from Module B (SourceEval)'s EvaluateGroupCellSource and Module C (TriggerPush)'s
' PushCellDerivedValuesToMembers.
Public Sub ReportMultipleTriggers()
    On Error Resume Next
    If Not mbMultiShown Then
        LangManager.ShowStatusT "CalculationMultipleTriggers"
        mbMultiShown = True
    End If
End Sub

' Own key, not a reuse of ReportMultipleTriggers: that message names "cells", but a Lvl*-collision may
' involve no cell at all. Public: called from Module B (SourceEval) and Module C (TriggerPush).
Public Sub ReportMultipleLvlTriggers()
    On Error Resume Next
    If Not mbMultiLvlShown Then
        LangManager.ShowStatusT "CalculationMultipleLvlTriggers"
        mbMultiLvlShown = True
    End If
End Sub

' Own key: this ambiguity is between LINKED elements, not cells or levels by name. Public: called from
' Module B (SourceEval)'s EvaluateGroupColor.
Public Sub ReportMultipleColorCandidates()
    On Error Resume Next
    If Not mbMultiColorShown Then
        LangManager.ShowStatusT "CalculationMultipleColorCandidates"
        mbMultiColorShown = True
    End If
End Sub

' Own key: the ambiguity is between measurable GEOMETRIES (first in scan order wins), not trigger cells.
' Public: called from Module B (SourceEval)'s EvaluateGroupLength.
Public Sub ReportMultipleGeometries()
    On Error Resume Next
    If Not mbMultiGeoShown Then
        LangManager.ShowStatusT "CalculationMultipleGeometries"
        mbMultiGeoShown = True
    End If
End Sub

' Safe "array has at least one element" check (mirrors ElementChangeHandler.HasElements). UBound
' returns -1 for an empty array and raises for an uninitialised one. Public: called from Module B
' (SourceEval)'s group scans.
Public Function HasElements(ByRef arr() As element) As Boolean
    On Error Resume Next
    HasElements = False
    If UBound(arr) <> -1 Then HasElements = True
    On Error GoTo 0
End Function
