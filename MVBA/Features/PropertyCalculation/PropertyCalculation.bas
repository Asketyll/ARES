' Module: PropertyCalculation
' Description: The VALUE-CALCULATION engine for ARES custom properties (redecoupage epic 11, AWAKE in
'              phase 2 - epic 14). It computes each custom property's VALUE from a per-property CALC RULE
'              and writes it - but ONLY where that property is ALREADY ATTACHED (by a PropertyTagging
'              rule). It NEVER attaches, and never calls CustomPropertyHandler attach/detach directly. A
'              member that does not carry the target property is SKIPPED (the frontier: attach/detach is
'              the tagger's domain). Opt-in, OFF by default (ARES_Property_Calc).
'
'              CALC GRAMMAR (ARES_Calc_Rules, epic 14):  "rule ; rule ; ..."  where each rule is
'                  Prop[name] [& condition]* = Source
'                - Prop[name] = the TARGET property (left-most). One name, NO wildcard (a target is named,
'                    not matched). Syntactic-only validation; membership in the ARES DGNLib is NOT
'                    hard-checked (the runtime frontier is the guard - the engine only writes where
'                    the property is actually attached).
'                - [& condition]* = OPTIONAL conditions, the SAME grammar as the tag rules (Lvl/Cell/Type,
'                    &, !, */? wildcards), delegated verbatim to the shared RuleGrammar module.
'                - Source (right of "="), keyword + optional [arg], arity-checked, unknown rejected. Two
'                    families: a GROUP source (keyword prefixed "Cell"/"Group") scans the element's graphic
'                    group INCLUDING itself for a matching member and reads off THAT member; a SELF source
'                    reads the bearing element's own attribute.
'                    * CellText[pattern] - the full text (StringsInEl.GetConcatenatedText) of a cell in the
'                        element's graphic group whose name matches pattern (wildcards OK; pattern may be
'                        several ARES_VAR_DELIMITER ("|") - separated alternatives, e.g. "ASUF*|SP0*|Bois*" -
'                        the SAME alternation as a tag/calc CONDITION's Cell[name|name|...]). INCLUDES the
'                        bearing element itself (a matching cell yields its own text; an ungrouped matching
'                        cell is a group of one). A GROUP source (driven by the matching cell).
'                    * CellCoord[pattern] / CellId[pattern] / CellLvl[pattern] / CellColor[pattern] /
'                        CellStyle[pattern] / CellWeight[pattern] - same matching-cell scan as CellText
'                        (self-included, same pattern/alternation grammar), reading that SAME cell's "X;Y"
'                        coordinates (anchor cascade, ARES_Round decimals, no [n] override) / DLongToString
'                        ID / level name / color index / line-style name / line-weight of that MATCHING
'                        CELL, not the bearing element's own - use these to read the tagging cell's own
'                        attributes on a member it pushed properties to (e.g.
'                        Prop[Coordonnee]=CellCoord[ASUF*|SP0*|Bois*|BO *|APO*|AH6*] where Coord alone would
'                        only ever yield the bearing element's own position). All GROUP sources except
'                        CellId are stable-pushed: a change on the matching cell (text/coord/level/color/
'                        style/weight) is re-pushed to the OTHER members by the trigger-cell pass (see
'                        below); CellId never needs a push (an ID cannot change).
'                    * Value[text] - a fixed literal value (Value[] empty = invalid). A SELF source.
'                    * Coord / Coord[n] - the "X;Y" coordinates of the bearing element (n = decimals,
'                        default = ARES_Round), via a deterministic anchor cascade. A SELF source
'                        (recomputes on Modify). Coordinates are ALREADY master units - no UOR scaling.
'                    * Id - the bearing element's ID via DLongToString (mandatory DLong helper). A SELF
'                        source (stable).
'                    * Lvl / Color / Style / Weight - the bearing element's OWN level name / color index /
'                        line-style name / line-weight (guarded by IsGraphical; "" when not graphical or
'                        when the specific attribute is unavailable - e.g. a level-less cell header - never
'                        a fabricated value). SELF sources (recompute on Modify).
'                    * Length / Length[n] - the bearing element's OWN geometry length (Length.GetLength; n =
'                        decimals, default = ARES_Length_Round), ONLY when the bearing element itself is a
'                        length-capable type (Line/Arc/Shape/ComplexShape/ComplexString); "" otherwise (no
'                        silent 0). A SELF source - use this when the calc-target property is attached
'                        directly to the geometry.
'                    * GroupLength / GroupLength[n] - same decimals as Length, but reads the FIRST length-
'                        capable element found scanning the graphic group INCLUDING itself (scan order, no
'                        name pattern - geometry has no name to match against, unlike Cell* sources); "" when
'                        the group (or the ungrouped bearing element itself) holds no length-capable member.
'                        A GROUP source - the common case: the calc-target property sits on a TEXT/CELL and
'                        the length comes from the linked line/arc/shape (mirrors the Auto Lengths pattern).
'                - Several rules for the SAME property: the FIRST rule that MATCHES wins (order = priority;
'                    put specific rules before general ones).
'              Example:  Prop[Repere]&Cell[ETIREF]=Value[REF] ; Prop[Repere]=CellText[ETI*] ; Prop[XY]=Coord ;
'                        Prop[Coordonnee]=CellCoord[ASUF*|SP0*] ; Prop[Longueur]=GroupLength[1]
'
'              ONE bracket-depth-aware parser (ParseCalcRule) is the single source of truth: it splits on
'              the depth-0 "=" (RuleGrammar.FindTopLevelChar), the LEFT side on the depth-0 "&"
'              (RuleGrammar.SplitTopLevel) - segment[0] MUST be Prop[name], segments[1..] are conditions
'              (RuleGrammar.ParseCondition) - and the RIGHT side into a single arity-checked Source.
'              EnsureCalcRulesParsed caches (skip a bad rule fail-closed); RefreshCalcRules invalidates.
'              ValidateAndNormalizeCalcRule(sRule, sCanonical) is the read-only validate-AND-normalise the
'              options editor (14-3) calls on commit: "" + COMPACT canonical form on valid, a targeted
'              reason on invalid. CalcRuleHasNoEffect(sRule, segments) flags a dead CONDITION combo (via
'              RuleGrammar.ConditionsHaveContradiction, ignoring the Prop target) for the preview.
'
'              ENGINE WAKE - two Depth-0 passes, both routing every write through ApplyValueToSibling
'              (its frontier + compare-guard + non-empty->empty transition guard + delegated
'              detach-on-empty are load-bearing loop-safety BLOCKERs):
'                - BEARING pass: for each DISTINCT calc-target property P that oEl carries, resolve its
'                    value from the FIRST matching calc rule and write it (compare-guarded). This is the
'                    single code path for "fill on attach", "recompute Coord on move", "re-pull CellText",
'                    and "reconcile on a neighbour's delete" (the surviving matching cell's text, or "" ->
'                    transition-guarded clear/detach when none survives). A property with no matching rule
'                    is LEFT UNTOUCHED (the engine only governs what a rule matches).
'                - TRIGGER-CELL pass: when oEl is a trigger cell (its name matches ANY pushable GROUP
'                    source's pattern - CellText/CellCoord/CellLvl/CellColor/CellStyle/CellWeight), push the
'                    freshly-read attribute to the OTHER group members carrying that rule's target - the
'                    members MicroStation did NOT re-queue when only the cell's own attribute changed (a cell
'                    is a Branch-1/text-cell element in ElementChangeHandler; only a Branch-2/geometric change
'                    auto-re-queues every group member). First-match guarded (a member whose P is governed by
'                    an earlier rule is left alone). CellId never needs a push (an ID cannot change, so the
'                    bearing pass's own computation on the member stays correct forever). GroupLength needs NO
'                    push of its own: its source is a GEOMETRY element, and ElementChangeHandler's Branch 2
'                    reacts when that geometry changes - gated on ARES_Update_Lengths OR
'                    PropertyCalculation.HasGroupLengthRules (a GroupLength rule pulls its own trigger into
'                    that gate so it works even with Auto Lengths OFF). Branch 2's own recursive walk runs at
'                    Depth > 0, which SKIPS this module's Depth-0-only hook, and its queue-add is wiped by
'                    IdleEventHandler.CleanupHandler before a fresh Depth-0 pass ever reaches the sibling - so
'                    Branch 2 also calls PropertyCalculation.ProcessElement on each linked sibling DIRECTLY
'                    (inline, not via the queue), exactly mirroring how Auto Lengths' own trigger-text work
'                    already runs inline there regardless of Depth.
'
'              Deletion is reconciled by the BEARING pass on the members ShouldQueueForDeletion already
'              re-queues (Link.GetLink(BeforeChange)) - no pending-clear machinery (retired in 14-2).
'
'              PropertyRendering (epic 15) is a READER of this module, in three narrow places: a value
'              STATE CHANGE notes the sibling for the renderer's bounded repaint hop; the CellText source
'              asks the renderer which sub-texts to EXCLUDE, so a rendered value can never feed the value
'              that governs it; and GetCalcRuleForProperty answers "where does this value come from" at
'              bind time. This module still writes values only - it knows nothing about text.
'
'              Emptying semantics: when a CellText source has no surviving matching cell it yields "";
'              ApplyValueToSibling then CLEARS the value (property stays attached) by default, or with
'              ARES_Calc_Detach_Empty ON DELEGATES the detach to the tagger (PropertyTagging.
'              DetachRuleProperty) - the ONLY detach path, gated on the non-empty->empty transition so a
'              re-attaching rule cannot oscillate it (termination).
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (global ARESConfig), RuleGrammar, CustomPropertyHandler,
'               PropertyTagging, PropertyRendering, StringsInEl, Link, Length, LangManager,
'               ErrorHandlerClass (global ErrorHandler)

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
' Runtime clamp so VBA Round never faults on an absurd decimal count (no coordinate/length ever needs > 15
' places).
Private Const SOURCE_ROUND_CLAMP As Long = 15

' Source vocabulary of a calc rule's right-hand side (canonicalised, case-insensitive on input). csCell*
' are GROUP sources (matching-cell scan, self-included); the rest are SELF sources (the bearing element's
' own attribute) except csGroupLength (GROUP, scanned by TYPE not name pattern).
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
    csLength
    csGroupLength
End Enum

' One parsed calc rule: Prop[TargetProp] [& conditions]* = Source. Conditions() (RuleGrammar.RuleCondition)
' is bounded by nCond. SourceArg holds the pattern (CellText/CellCoord/CellId), the fixed text (Value), the
' decimals string (Coord[n]) - empty for Id and bare Coord.
Private Type CalcRuleInfo
    TargetProp As String
    Conditions() As RuleGrammar.RuleCondition
    nCond As Long
    SourceKind As CalcSource
    SourceArg As String
End Type

Private mCalcRules() As CalcRuleInfo
Private mnCalcCount As Long
Private mbCalcParsed As Boolean

' One-shot guards so the calculation statuses (CalculationValueRejected / CalculationNoTarget /
' CalculationMultipleTriggers) each surface only once per PROCESSED ELEMENT. Reset at the start of each
' ProcessElement; the fault one (Rejected) also keeps its English log on every occurrence; NoTarget and
' Multiple are user feedback (status-only, no log).
Private mbRejectedShown As Boolean
Private mbNoTargetShown As Boolean
Private mbMultiShown As Boolean

'######################################################################################################################
'                                          PUBLIC SURFACE
'######################################################################################################################

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

' Read-only validate-AND-normalise for ONE calc rule (the seam the 14-3 editor writes through). Returns:
'   - "" with sCanonical = "" when the rule is empty (the caller treats it as a delete);
'   - "" with sCanonical = the COMPACT canonical stored form when the rule is valid;
'   - a short English reason (fault/log channel) when the rule is invalid.
' It calls the SAME ParseCalcRule the runtime parser uses, so it accepts exactly what the parser accepts
' (no drift). Syntactic only - no DGNLib membership check on Prop[name] (the runtime frontier is the guard).
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

' Trigger test (re-wired AWAKE, epic 14; extended for CellCoord/CellLvl/CellColor/CellStyle/CellWeight). A
' trigger cell is a CELL, in a REAL graphic group, whose name matches a PUSHABLE Cell* source's pattern
' (IsPushableCellSourceKind) of at least one calc rule. Drives the trigger-cell pass (pushing a changed
' cell's attribute to the members MicroStation did not re-queue). CellId is excluded (an ID never changes,
' so it never needs a push). An ungrouped matching cell is NOT a trigger (it has no other members; its own
' value is handled by the bearing pass via each GROUP source's self-inclusion).
Public Function IsTriggerCell(ByVal oEl As element) As Boolean
    On Error GoTo ErrorHandler

    IsTriggerCell = False
    If oEl Is Nothing Then Exit Function
    If Not oEl.IsCellElement Then Exit Function
    If oEl.GraphicGroup = ARES_DEFAULT_GRAPHIC_GROUP_ID Then Exit Function

    EnsureCalcRulesParsed
    If mnCalcCount = 0 Then Exit Function

    IsTriggerCell = AnyPushableSourcePatternMatches(oEl.AsCellElement.Name)
    Exit Function

ErrorHandler:
    IsTriggerCell = False
End Function

' Depth-0 hook, called from ElementChangeHandler.ProcessElement (before the graphic-group filter) when the
' feature is enabled. Runs the two passes:
'   (1) BEARING pass - fill/recompute each calc-target property oEl carries from its first matching rule.
'   (2) TRIGGER-CELL pass - if oEl is a trigger cell, push its text to the OTHER members carrying the fed
'       target (the members not otherwise re-queued when only the cell's text changed).
' Every write routes through ApplyValueToSibling (frontier + compare-guard + transition guard + delegated
' detach). The one-shot status guards are reset here (per processed element).
Public Sub ProcessElement(ByVal oEl As element)
    On Error GoTo ErrorHandler

    mbRejectedShown = False
    mbNoTargetShown = False
    mbMultiShown = False

    If oEl Is Nothing Then Exit Sub

    EnsureCalcRulesParsed
    If mnCalcCount = 0 Then Exit Sub

    BearingPass oEl

    If IsTriggerCell(oEl) Then PushCellDerivedValuesToMembers oEl
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation.ProcessElement"
End Sub

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
                ParseSource = "empty source (expected CellText/CellCoord/CellId/CellLvl/CellColor/CellStyle/CellWeight/Value/Coord/Id/Lvl/Color/Style/Weight/Length/GroupLength)"
            Else
                ParseSource = "unknown source '" & kw & "' (expected CellText/CellCoord/CellId/CellLvl/CellColor/CellStyle/CellWeight/Value/Coord/Id/Lvl/Color/Style/Weight/Length/GroupLength)"
            End If
            Exit Function
    End Select
    Exit Function

ErrorHandler:
    ParseSource = "invalid source"
End Function

' Shared arity check for every GROUP source that takes a mandatory non-empty [pattern] (CellText/CellCoord/
' CellId/CellLvl/CellColor/CellStyle/CellWeight). Returns False + sReason on a missing/empty pattern, True
' + outPat (trimmed) on success.
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

' Resolve P's value on oEl: find the FIRST calc rule whose target is P and whose conditions match oEl, then
' evaluate its source. bHasRule distinguishes "no rule governs P here" (False -> the caller leaves P alone)
' from "a rule yields the empty string" (True with "" -> the caller clears/detaches, transition-guarded).
' READ-ONLY (no attach, no write, no frontier): computes "what value would P get on this element". Public
' so PropertyCalculationTest can assert first-match + all four sources DGNLib-free (and a future preview
' could reuse it), mirroring the phase-1 read-only test seam.
Public Function ResolvePropertyValue(ByVal P As String, ByVal oEl As element, ByRef bHasRule As Boolean) As String
    On Error GoTo ErrorHandler

    ResolvePropertyValue = ""
    bHasRule = False

    Dim idx As Long
    idx = FindCalcRuleForProperty(P, oEl)
    If idx < 0 Then Exit Function

    bHasRule = True
    ResolvePropertyValue = EvaluateSource(mCalcRules(idx), oEl)
    Exit Function

ErrorHandler:
    ' Fail-closed: withhold a value AND signal "no rule" so the caller does not clear on a fault.
    ResolvePropertyValue = ""
    bHasRule = False
End Function

' Read-only "where does P's value come from on this element" seam. Returns False when no calc rule
' governs P here; otherwise True with the source kind, its bracket argument and a canonical
' "Source[arg]" string. Public because PropertyRendering needs it at BIND time to tell the user which
' engine owns the value they are about to display, and to spot the one static cycle v1 can detect (a
' token rendered inside the very cell whose text feeds the property). Same ByRef-out shape as
' ResolvePropertyValue's bHasRule. CalcRuleInfo stays a Private Type - only scalars cross the boundary.
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
' conditions match oEl; -1 if none. Order in the config = priority (specific rules first).
Private Function FindCalcRuleForProperty(ByVal P As String, ByVal oEl As element) As Long
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

' Evaluate a calc rule's Source against the bearing element. Returns the computed/fixed string ("" when a
' CellText/GroupLength source finds no matching member, or a SELF attribute is unavailable). Coordinates are
' ALREADY master units (mvba-docs) - no scaling.
Private Function EvaluateSource(ByRef r As CalcRuleInfo, ByVal oEl As element) As String
    On Error GoTo ErrorHandler

    EvaluateSource = ""
    Select Case r.SourceKind
        Case csCellText, csCellCoord, csCellId, csCellLvl, csCellColor, csCellStyle, csCellWeight
            EvaluateSource = EvaluateGroupCellSource(oEl, r.SourceArg, r.SourceKind)
        Case csValue
            EvaluateSource = r.SourceArg
        Case csCoord
            Dim pt As Point3d
            If GetElementAnchorPoint(oEl, pt) Then
                EvaluateSource = FormatCoord(pt, ResolveDecimals(r.SourceArg, GetCoordDefaultDecimals()))
            Else
                ' No valid anchor (a geometry fault, or a non-graphical bearing element) -> yield NO value
                ' (never a fabricated "0;0"), the same "no value rather than a wrong value" philosophy as a
                ' FormatCoord fault. Log the technical anomaly (English, Number 0); ResolvePropertyValue then
                ' returns "" -> the transition-guarded clear (safe).
                ErrorHandler.HandleError "Property calculation: no anchor point for Coord source", 0, "", "PropertyCalculation.EvaluateSource"
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

' Shared GROUP scan for every Cell* source: scan the bearing element's graphic group INCLUDING itself
' (Link.GetLink ReturnMe:=True) and return the FIRST cell (scan order) whose name matches sPattern via
' foundCell (Nothing when none); nMatch = total matching-cell count, left to each caller to act on (>= 2
' drives the one-shot multi-trigger warning). For an UNGROUPED bearing element Link.GetLink returns nothing,
' so the element is its own sole candidate (a group of one).
Private Function FindFirstMatchingCellInGroup(ByVal oEl As element, ByVal sPattern As String, ByRef foundCell As element, ByRef nMatch As Long) As Boolean
    On Error GoTo ErrorHandler

    FindFirstMatchingCellInGroup = False
    Set foundCell = Nothing
    nMatch = 0

    Dim cands() As element
    cands = Link.GetLink(oEl, True)

    If HasElements(cands) Then
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
' of WHICH attribute is being read off the matching cell.
Private Function EvaluateGroupCellSource(ByVal oEl As element, ByVal sPattern As String, ByVal kind As CalcSource) As String
    On Error GoTo ErrorHandler

    EvaluateGroupCellSource = ""

    Dim foundCell As element
    Dim nMatch As Long
    If FindFirstMatchingCellInGroup(oEl, sPattern, foundCell, nMatch) Then
        EvaluateGroupCellSource = ReadCellSourceValue(foundCell, kind)
    End If
    If nMatch >= 2 Then ReportMultipleTriggers
    Exit Function

ErrorHandler:
    EvaluateGroupCellSource = ""
End Function

' Read ONE attribute off a SPECIFIC cell element (already located - either by FindFirstMatchingCellInGroup
' during the bearing pass, or as the trigger cell itself during the push pass). Never fabricates a value:
' a missing Level/LineStyle yields "" (mirrors the no-anchor Coord/CellCoord philosophy). Coordinates use
' the default decimals (no [n] override on a Cell* source - the bracket already carries the pattern).
Private Function ReadCellSourceValue(ByVal oCell As element, ByVal kind As CalcSource) As String
    On Error GoTo ErrorHandler

    ReadCellSourceValue = ""
    Select Case kind
        Case csCellText
            ' Containment (epic 15): a sub-text the RENDERER writes must not feed the value that governs
            ' it, or the value would ratchet on its own output. The exclusion is a function of the SOURCE
            ' CELL, not of the rule, so it is resolved HERE rather than threaded through six signatures -
            ' and that is also what keeps PushCellDerivedValuesToMembers' lazy cache correct: the cache is
            ' keyed on the CalcSource ordinal alone, which stays sound only because within one push pass
            ' there is exactly ONE source cell and therefore exactly one exclusion set. A future
            ' RULE-dependent exclusion would break that cache and must re-key it.
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
                ReadCellSourceValue = FormatCoord(pt, GetCoordDefaultDecimals())
            Else
                ErrorHandler.HandleError "Property calculation: no anchor point for CellCoord source", 0, "", "PropertyCalculation.ReadCellSourceValue"
            End If
        Case csCellId
            ReadCellSourceValue = DLongToString(oCell.ID)
        Case csCellLvl
            If Not oCell.Level Is Nothing Then ReadCellSourceValue = oCell.Level.Name
        Case csCellColor
            ReadCellSourceValue = CStr(oCell.Color)
        Case csCellStyle
            If Not oCell.LineStyle Is Nothing Then ReadCellSourceValue = oCell.LineStyle.Name
        Case csCellWeight
            ReadCellSourceValue = CStr(oCell.LineWeight)
    End Select
    Exit Function

ErrorHandler:
    ReadCellSourceValue = ""
End Function

' True when sName matches at least one of sPattern's ARES_VAR_DELIMITER ("|") - separated parts
' (case-insensitive, wildcards via RuleGrammar.LikeCI) - the SAME multi-prefix alternation as the tag/calc
' CONDITION grammar (Cell[ASUF*|SP0*|...]), so a CellText/CellCoord/CellId[pattern] source can feed off the
' same cell-name families a Property Tagging @ rule already groups. A single pattern with no "|" is just a
' one-element split (behavior-identical to the pre-alternation scalar match). Shared by every cell-name-vs-
' pattern comparison in this module (IsMatchingCell, the trigger-cell detector, the push loop) so the
' alternation is honoured consistently everywhere, not just on the initial bearing-pass computation.
Private Function MatchesAnyPattern(ByVal sName As String, ByVal sPattern As String) As Boolean
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

'######################################################################################################################
'                                          ENGINE - TRIGGER-CELL PASS
'######################################################################################################################

' True for the Cell* source kinds that are STABLE-PUSHED (CellText/CellCoord/CellLvl/CellColor/CellStyle/
' CellWeight): a change on the matching cell can leave its group siblings un-re-queued (a cell is a
' Branch-1/text-cell element in ElementChangeHandler; no automatic group-wide re-queue), so the trigger-cell
' pass must push. CellId is excluded - an ID can never change, so it never needs a push (see IsTriggerCell).
Private Function IsPushableCellSourceKind(ByVal kind As CalcSource) As Boolean
    IsPushableCellSourceKind = (kind = csCellText Or kind = csCellCoord Or kind = csCellLvl Or _
                                 kind = csCellColor Or kind = csCellStyle Or kind = csCellWeight)
End Function

' Trigger-cell pass: oCell is a trigger cell whose text/coordinates/level/color/style/weight may have
' changed while its group members were NOT re-queued. For each OTHER group member M carrying a property P
' fed by a pushable Cell* rule matching oCell's name, and where THIS rule is M's first-match for P (AC3 - a
' member governed by an earlier rule for P is left alone), push the rule-appropriate attribute (compare-
' guarded via ApplyValueToSibling) read off oCell via ReadCellSourceValue - each SourceKind computed at most
' ONCE per call (lazy, cached by enum ordinal, shared across every matching rule/member). Discoverability:
' matching members exist but none carry a fed target -> one-shot CalculationNoTarget. Two competing cells
' for one P -> one-shot CalculationMultipleTriggers (last-processed wins).
Private Sub PushCellDerivedValuesToMembers(ByVal oCell As element)
    On Error GoTo ErrorHandler

    Dim members() As element
    members = Link.GetLink(oCell)                 ' OTHER members (self handled by the bearing pass)
    If Not HasElements(members) Then Exit Sub

    Dim sName As String
    sName = oCell.AsCellElement.Name

    ' Lazy per-SourceKind cache, indexed by the CalcSource enum ordinal (csGroupLength is the highest).
    Dim cacheVal(0 To csGroupLength) As String
    Dim cacheReady(0 To csGroupLength) As Boolean

    Dim i As Long, ri As Long, kIdx As Long
    Dim m As element
    Dim P As String
    Dim nCarried As Long
    nCarried = 0
    For i = LBound(members) To UBound(members)
        Set m = members(i)
        If Not m Is Nothing Then
            For ri = 0 To mnCalcCount - 1
                If IsPushableCellSourceKind(mCalcRules(ri).SourceKind) Then
                    If MatchesAnyPattern(sName, mCalcRules(ri).SourceArg) Then
                        P = mCalcRules(ri).TargetProp
                        If CustomPropertyHandler.IsItemAttachedToElement(m, P) Then
                            nCarried = nCarried + 1
                            ' First-match guard (AC3): push only where THIS rule governs m's P.
                            If FindCalcRuleForProperty(P, m) = ri Then
                                kIdx = CLng(mCalcRules(ri).SourceKind)
                                If Not cacheReady(kIdx) Then
                                    cacheVal(kIdx) = ReadCellSourceValue(oCell, mCalcRules(ri).SourceKind)
                                    cacheReady(kIdx) = True
                                End If
                                ApplyValueToSibling m, P, cacheVal(kIdx)
                            End If
                        End If
                    End If
                End If
            Next ri
        End If
    Next i

    ' Discoverability: siblings match a rule but NONE carry its target -> attach never happened.
    If nCarried = 0 Then ReportNoTarget

    ' Multi-trigger: another cell in the group also feeds one of the pushed targets (last-processed wins).
    If GroupHasCompetingTrigger(oCell) Then ReportMultipleTriggers
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation.PushCellDerivedValuesToMembers"
End Sub

' True when oCell's graphic group holds at least one OTHER cell that also matches some pushable Cell*
' source's pattern (i.e. two label/anchor cells could feed the same target) - the multi-trigger condition.
' Read-only.
Private Function GroupHasCompetingTrigger(ByVal oCell As element) As Boolean
    On Error GoTo ErrorHandler

    GroupHasCompetingTrigger = False

    Dim members() As element
    members = Link.GetLink(oCell)                 ' OTHER members only
    If Not HasElements(members) Then Exit Function

    Dim i As Long
    For i = LBound(members) To UBound(members)
        If Not members(i) Is Nothing Then
            If members(i).IsCellElement Then
                If AnyPushableSourcePatternMatches(members(i).AsCellElement.Name) Then
                    GroupHasCompetingTrigger = True
                    Exit Function
                End If
            End If
        End If
    Next i
    Exit Function

ErrorHandler:
    GroupHasCompetingTrigger = False
End Function

' True when sName matches a pushable Cell* source's pattern of at least one calc rule (assumes the cache is
' parsed). CellId is excluded - it is never pushed (see IsPushableCellSourceKind/IsTriggerCell), so a cell
' that only feeds a CellId rule must not be treated as a trigger.
Private Function AnyPushableSourcePatternMatches(ByVal sName As String) As Boolean
    On Error GoTo ErrorHandler

    AnyPushableSourcePatternMatches = False
    Dim i As Long
    For i = 0 To mnCalcCount - 1
        If IsPushableCellSourceKind(mCalcRules(i).SourceKind) Then
            If MatchesAnyPattern(sName, mCalcRules(i).SourceArg) Then
                AnyPushableSourcePatternMatches = True
                Exit Function
            End If
        End If
    Next i
    Exit Function

ErrorHandler:
    AnyPushableSourcePatternMatches = False
End Function

'######################################################################################################################
'                                          GEOMETRY - Coord ANCHOR CASCADE
'######################################################################################################################

' Deterministic anchor point of an element, for the Coord/CellCoord sources (the latter passes the MATCHING
' CELL, not the bearing element). Returns True and fills pt when an anchor is available; returns False ONLY
' when even the universal Range-centre seed cannot be computed (a non-graphical element, or a Range fault) -
' the caller then yields "" + logs, NEVER a fabricated
' coordinate. It NEVER returns a fabricated (0,0,0): the Range centre ((Low+High)/2) is seeded first, and a
' type-specific anchor OVERRIDES it only on success, so a per-branch geometry fault (e.g. the
' applicability-unverified AsClosedElement.Centroid raising on a closed element) degrades to the Range
' centre, never to the origin. Coordinates are already master units (no UOR scaling).
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

' Attempt the type-specific anchor for the Coord source, isolated so a raise NEVER reaches the caller's
' Range-centre seed. Each As*/geometry API is verified in mvba-docs (a review BLOCKER to use an unverified
' signature): cell -> AsCellElement.Origin; shared cell -> AsSharedCellElement.Origin; text ->
' AsTextElement.Origin (the User Origin, per mvba-docs); text node -> AsTextNodeElement.Origin; line ->
' AsLineElement.Origin; arc -> AsArcElement.CenterPoint; ellipse -> AsEllipseElement.CenterPoint; closed
' (Shape/ComplexShape) -> AsClosedElement.Centroid (a method returning Point3d, no ByRef args; its
' applicability to a closed element is UNDOCUMENTED, so a raise here degrades to the Range-centre seed).
' Returns False (-> the caller keeps the seed) when the element has no specific anchor OR the read faults.
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

' Format a point as "X;Y" with dec decimals, mirroring Auto_Lengths.cls CStr(Round(...)). The decimal
' separator is locale-dependent (a comma-decimal locale yields "1,5;2,5" - the ";" field separator does
' not collide). dec is clamped so VBA Round never faults on an absurd count.
Private Function FormatCoord(ByRef pt As Point3d, ByVal dec As Long) As String
    On Error GoTo ErrorHandler

    Dim d As Long
    d = dec
    If d < 0 Then d = 0
    If d > SOURCE_ROUND_CLAMP Then d = SOURCE_ROUND_CLAMP

    FormatCoord = CStr(Round(pt.X, d)) & ";" & CStr(Round(pt.Y, d))
    Exit Function

ErrorHandler:
    FormatCoord = ""
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

' Shared GROUP scan for GroupLength: scan the bearing element's graphic group INCLUDING itself (Link.GetLink
' ReturnMe:=True) and return the FIRST length-capable member (scan order, NOT a name pattern - geometry has
' no name to match, unlike a Cell* source) via foundGeo (Nothing when none); nMatch = total length-capable
' member count. For an UNGROUPED bearing element Link.GetLink returns nothing, so the element is its own sole
' candidate (a group of one) - mirrors FindFirstMatchingCellInGroup.
Private Function FindFirstLengthCapableInGroup(ByVal oEl As element, ByRef foundGeo As element, ByRef nMatch As Long) As Boolean
    On Error GoTo ErrorHandler

    FindFirstLengthCapableInGroup = False
    Set foundGeo = Nothing
    nMatch = 0

    Dim cands() As element
    cands = Link.GetLink(oEl, True)

    If HasElements(cands) Then
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

' GroupLength evaluation: the length of the FIRST length-capable element found scanning the group
' (self-included, deterministic scan order - the same first-match philosophy as calc rules themselves); ""
' when neither the group nor the ungrouped bearing element itself holds one. Unlike the Cell* sources this
' never warns on multiple candidates (a group commonly holds several geometric members - e.g. an outline
' plus inner detail - without that being a misconfiguration the way two competing label cells would be).
Private Function EvaluateGroupLength(ByVal oEl As element, ByVal dec As Long) As String
    On Error GoTo ErrorHandler

    EvaluateGroupLength = ""

    Dim foundGeo As element
    Dim nMatch As Long
    If FindFirstLengthCapableInGroup(oEl, foundGeo, nMatch) Then
        EvaluateGroupLength = EvaluateOwnLength(foundGeo, dec)
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

' The frontier + compare-before-write on a single sibling (loop-safety). Returns True when s ALREADY
' carries the target property P (whether or not a write happened) - the caller counts these to detect
' the "no member carries P" misconfiguration. The value engine NEVER attaches and never calls
' CustomPropertyHandler detach directly:
'   - P not attached (IsItemAttachedToElement False) -> SKIP (return False). Attach is the tagger's job.
'   - non-empty value, different from current       -> set (compare-guarded); rejection -> one-shot status.
'   - non-empty value, equal to current             -> no-op (loop-safety).
'   - empty value, current non-empty (a real emptying TRANSITION):
'         option OFF -> clear the value ("");  option ON -> delegate a detach to the tagger
'         (PropertyTagging.DetachRuleProperty). This is the ONLY detach path, gated on BOTH the option
'         AND the non-empty->empty transition (the load-bearing loop-safety guard).
'   - empty value, current already empty            -> no-op (transition guard: no re-detach).
' The read AND both writes address P's OWN item explicitly (ItemName:=P): with ItemName omitted, the
' handler resolves to the FIRST attached ARES item (Items.Find "*"), which misdirects the write as soon
' as the element carries a second calc property (the "Coord value lands in the wrong property" bug).
Private Function ApplyValueToSibling(ByVal s As element, ByVal P As String, ByVal value As String) As Boolean
    On Error GoTo ErrorHandler

    ApplyValueToSibling = False
    If s Is Nothing Then Exit Function

    ' Frontier: write only where P is ALREADY attached (HasItems, not Null-inference - an attached-but-
    ' empty property also reads back Null). Not attached -> skip; attach stays the tagger's domain.
    If Not CustomPropertyHandler.IsItemAttachedToElement(s, P) Then Exit Function
    ApplyValueToSibling = True

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
                ' Value STATE CHANGE: note the sibling for the renderer's bounded repaint hop. The
                ' pipeline does not re-queue a member whose value we just pushed, so without this its
                ' displayed text would stay stale until the element is touched again. A REJECTED write
                ' warrants no repaint. The element is passed as-is: the renderer resolves the graphic
                ' group itself behind its own IsGraphical guard, because reading .GraphicGroup here would
                ' RAISE on a non-graphical element and pollute this handler on every write.
                PropertyRendering.NoteDirtyGroup s
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
            ' The EMPTYING transition is noted too, whichever branch handled it: it is precisely what must
            ' re-materialise the literal token in a rendered text, on a sibling nothing else re-queues.
            ' Gated on the clear having been ACCEPTED, exactly like the non-empty branch above: a picklist
            ' or otherwise constrained property that rejects the clear leaves the value where it was, and a
            ' rejected write warrants no repaint. The detach branch has no return to test (DetachRuleProperty
            ' is a Sub), so it stays a transition.
            If bEmptied Then PropertyRendering.NoteDirtyGroup s
        End If
    End If
    Exit Function

ErrorHandler:
    ' A fault mid-write does not un-attach P; the return value (set True once past the frontier) is only
    ' used to detect "no member carried P", so leaving it as-is is correct.
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation.ApplyValueToSibling"
End Function

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

' Surface CalculationNoTarget ONCE per processed element: a trigger cell fired with siblings present but
' NONE carried its target property (the value engine writes only where a rule already attached P). USER
' FEEDBACK, not a fault - status-only, no English .log (like ReportMultipleTriggers). Hints the user to add
' an attach rule in Property Tagging.
Private Sub ReportNoTarget()
    On Error Resume Next
    If Not mbNoTargetShown Then
        LangManager.ShowStatusT "CalculationNoTarget"
        mbNoTargetShown = True
    End If
End Sub

' Surface CalculationMultipleTriggers ONCE per processed element (deduped via mbMultiShown, reset in
' ProcessElement). USER FEEDBACK, not a fault: per the Design Note it is status-only and does NOT write an
' English .log line (unlike ReportRejected). Last-processed wins.
Private Sub ReportMultipleTriggers()
    On Error Resume Next
    If Not mbMultiShown Then
        LangManager.ShowStatusT "CalculationMultipleTriggers"
        mbMultiShown = True
    End If
End Sub

' Safe "array has at least one element" check (mirrors ElementChangeHandler.HasElements). UBound
' returns -1 for an empty array and raises for an uninitialised one.
Private Function HasElements(ByRef arr() As element) As Boolean
    On Error Resume Next
    HasElements = False
    If UBound(arr) <> -1 Then HasElements = True
    On Error GoTo 0
End Function
