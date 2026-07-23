' Module: PropertyTagging
' Description: The SOLE attach/detach engine for ARES custom properties. Auto-attaches ARES custom
'              properties to elements as they are created / modified, driven by configurable rules
'              (ARES_Property_Rules). Attach-only on match: the property appears (empty) on the element
'              for the user to fill / pick from its native value-list dropdown.
'
'              Called from ElementChangeHandler.ProcessElement (deferred, on idle) when
'              ARES_Auto_Properties is True. Rules are parsed once and cached - call RefreshRules
'              after changing ARES_Property_Rules at runtime.
'
'              GRAMMAR v2 (ARES_Property_Rules):  "rule ; rule ; ..."  where each rule is
'                  [@] condition [& condition]* = prop[|prop]*
'                - condition = [!] Keyword[name|name|...]
'                    Keyword (case-insensitive, stored canonical): Lvl (level name), Cell (cell name -
'                    IMPLIES the element is a cell), Type (element type name via
'                    MicroStationDefinition.StringToMsdElementType, plus the special token Cell = any cell).
'                    An unknown keyword rejects the whole rule (reserves the namespace, fail-closed).
'                - & = AND between conditions. OR between families = several rules (";"-separated).
'                - ! = strict negation of a condition ("!Cell[A]" = "is NOT a cell named A"; a line
'                    satisfies it). "a cell but not A" = "Type[Cell]&!Cell[A]".
'                - * (any run) / ? (any single char) = wildcards on Lvl / Cell names (VBA Like, # escaped,
'                    case-insensitive). Wildcards are NOT allowed in Type[...] (a type is an enum).
'                - @ = a RULE modifier (leading, normalised): the properties attach to the OTHER members
'                    of the matching element's graphic group (nothing without a real group). Without @,
'                    they attach to the matching element itself.
'                - Inside [...]: any literal EXCEPT the name separator "|" and the forbidden ";" / "[" /
'                    "]". So "=", "&", "@", "(", ")" are LITERAL inside brackets (Lvl[Poste=HTA], Lvl[R&D]).
'                - "(" and ")" are reserved at bracket depth 0 (rejected); literal inside [...].
'                - Right of "=": "|"-separated property names, everything literal ("@" literal); both
'                    sides of "=" must be non-empty. A prop containing "=" or ";" is rejected (the
'                    "|"-instead-of-";" mistake stays caught).
'              Example:  Type[Cell]&!Cell[A]=Repere ; @Cell[ETI0*]=Commune ; Lvl[WALLS]=Commune|Coupe_Type
'
'              ONE bracket-depth-aware parser (ParseOneRule) is the single source of truth: both
'              EnsureRulesParsed (skip fail-closed on any bad rule) and ValidateAndNormalizeRule call it,
'              so the validator accepts exactly what the parser accepts. v1 rules (level[:type]=prop,
'              @CellName=prop) have no recognised keyword => they are INVALID (skipped / refused); there
'              is no v1 parser, no migration, no bridge.
'
'              ValidateAndNormalizeRule(sRule, sCanonical) is the read-only validate-AND-normalise the
'              options form calls on every commit: "" + canonical form on a valid rule, a targeted reason
'              on an invalid one. RuleHasNoEffect(sRule, segments) is a read-only contradiction detector
'              (a syntactically valid rule that can never match) feeding the 13-3 coloured preview.
'
'              DetachRuleProperty(El, P) is the public detach service used by the (phase-1 DORMANT)
'              calculation engine's value-write scaffolding. The calculation engine no longer reads these
'              rules (grammar v2 removed the @cell=prop value seam; GetCellGroupProperties is gone).
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (global ARESConfig), CustomPropertyHandler, Link, MicroStationDefinition, ErrorHandlerClass (global ErrorHandler)

Option Explicit

Private Const RULE_SEPARATOR As String = ";"
Private Const SELECTOR_SEPARATOR As String = "="
Private Const COND_SEPARATOR As String = "&"
Private Const CELL_GROUP_MARKER As String = "@"
Private Const NEG_MARKER As String = "!"
Private Const BRK_OPEN As String = "["
Private Const BRK_CLOSE As String = "]"
Private Const PAREN_OPEN As String = "("
Private Const PAREN_CLOSE As String = ")"
' The name separator inside [...] and between property names is ARESConstants.ARES_VAR_DELIMITER ("|").

' Rule keyword vocabulary (canonicalised, case-insensitive on input).
Public Enum RuleKeyword
    rkLvl
    rkCell
    rkType
End Enum

' One parsed condition: [!] Keyword[name|name|...]. Names are kept VERBATIM (trimmed) for Like matching.
' For rkType, Names resolve to Types() (MsdElementType values) and/or MatchesAnyCell (the special "Cell"
' token = any cell, since there is no single MsdElementType for a cell and StringToMsdElementType("Cell")
' does not resolve to one).
Private Type RuleCondition
    Keyword As RuleKeyword
    Negated As Boolean
    Names() As String
    Types() As Long
    MatchesAnyCell As Boolean
End Type

' One parsed rule: [@] conditions (AND) = props. IsGroup is the "@" modifier (attach to the OTHER
' graphic-group members). nCond bounds the meaningful entries of Conditions().
Private Type RuleInfo
    IsGroup As Boolean
    Conditions() As RuleCondition
    nCond As Long
    Props() As String
End Type

Private mRules() As RuleInfo
Private mnRuleCount As Long
Private mbParsed As Boolean

' Force a re-parse of ARES_Property_Rules on the next match/apply (call after editing the variable).
Public Sub RefreshRules()
    mbParsed = False
End Sub

'######################################################################################################################
'                                          PUBLIC SURFACE
'######################################################################################################################

' True when at least one NON-group rule matches the element. Fast path for
' ElementChangeHandler.ShouldQueueElement: an ungrouped element cannot benefit from a "@" (group) rule
' (no other members), so only self-attach rules make it worth queueing. Keep the IsGraphical guard.
Public Function ElementMatchesAnyRule(ByVal oElement As element) As Boolean
    On Error GoTo ErrorHandler

    ElementMatchesAnyRule = False
    If oElement Is Nothing Then Exit Function

    EnsureRulesParsed
    If mnRuleCount = 0 Then Exit Function

    ' Non-graphical elements are never queued through this path.
    If Not oElement.IsGraphical Then Exit Function

    ' Resolve the level once (guarded). A cell header is graphical but has no Level; Cell/Type conditions
    ' must still be evaluated, so we do NOT exit on a missing level - we pass bHasLevel = False.
    Dim sLevel As String
    Dim bHasLevel As Boolean
    sLevel = ""
    bHasLevel = False
    If Not oElement.Level Is Nothing Then
        sLevel = oElement.Level.Name
        bHasLevel = True
    End If

    Dim i As Long
    For i = 0 To mnRuleCount - 1
        If Not mRules(i).IsGroup Then
            If RuleMatches(mRules(i), oElement, sLevel, bHasLevel) Then
                ElementMatchesAnyRule = True
                Exit Function
            End If
        End If
    Next i
    Exit Function

ErrorHandler:
    ElementMatchesAnyRule = False
End Function

' Attach the configured properties for every rule the element drives. For each matching rule: a "@"
' (group) rule fans the props out to each OTHER member of the element's graphic group (nothing without a
' real group); a plain rule attaches the props to the element itself. All attaches are idempotent
' (CustomPropertyHandler.AttachItemToElement is HasItems-guarded) -> loop-safe. Level is read once
' (guarded) and passed to the matcher, so Cell/Type rules still reach a level-less cell header.
Public Sub ApplyPropertyRules(ByVal oElement As element)
    On Error GoTo ErrorHandler

    If oElement Is Nothing Then Exit Sub

    EnsureRulesParsed
    If mnRuleCount = 0 Then Exit Sub

    Dim sLevel As String
    Dim bHasLevel As Boolean
    sLevel = ""
    bHasLevel = False
    If oElement.IsGraphical Then
        If Not oElement.Level Is Nothing Then
            sLevel = oElement.Level.Name
            bHasLevel = True
        End If
    End If

    Dim i As Long, j As Long
    For i = 0 To mnRuleCount - 1
        If RuleMatches(mRules(i), oElement, sLevel, bHasLevel) Then
            If mRules(i).IsGroup Then
                AttachGroupMembers oElement, mRules(i)
            Else
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

' Public detach service: remove a single property P from El. Called by the (phase-1 dormant) calculation
' engine's value-write scaffolding when it empties a value with ARES_Calc_Detach_Empty ON - detach is
' delegated here so ALL attach/detach stays inside PropertyTagging. Thin wrapper over
' CustomPropertyHandler.RemoveItemFromElement (itself HasItems-guarded, idempotent). Does NOT consult the
' parsed rules.
Public Sub DetachRuleProperty(ByVal El As element, ByVal P As String)
    On Error GoTo ErrorHandler

    If El Is Nothing Then Exit Sub
    If Len(Trim(P)) = 0 Then Exit Sub

    CustomPropertyHandler.RemoveItemFromElement El, P
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging.DetachRuleProperty"
End Sub

' Read-only validate-AND-normalise for ONE rule (the seam the editor writes through). Returns:
'   - "" with sCanonical = "" when the rule is empty (the caller treats it as a delete);
'   - "" with sCanonical = the canonical stored form when the rule is valid;
'   - a short English reason (fault/log channel) when the rule is invalid.
' It calls the SAME ParseOneRule the runtime parser uses, so it accepts exactly what the parser accepts
' (no drift). Canonical form is COMPACT (no spaces around "&"/"="; see RuleToCanonical). Syntactic only -
' no DGNLib membership check (a property may be authored later). Called from the options form on commit.
Public Function ValidateAndNormalizeRule(ByVal sRule As String, ByRef sCanonical As String) As String
    On Error GoTo ErrorHandler

    ValidateAndNormalizeRule = ""
    sCanonical = ""

    Dim s As String
    s = Trim(sRule)
    If Len(s) = 0 Then Exit Function

    Dim r As RuleInfo
    Dim sReason As String
    sReason = ParseOneRule(s, r)
    If Len(sReason) > 0 Then
        ValidateAndNormalizeRule = sReason
        Exit Function
    End If

    sCanonical = RuleToCanonical(r)
    Exit Function

ErrorHandler:
    ValidateAndNormalizeRule = "invalid rule"
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging.ValidateAndNormalizeRule"
End Function

' Read-only contradiction (dead-rule) detector on a SYNTACTICALLY VALID rule. Returns True (with the two
' conflicting condition segments, canonical text) when the rule can never match:
'   (a) two POSITIVE conditions of the same keyword with disjoint name-lists and NO wildcard among them
'       (Type[Line]&Type[Arc], Lvl[A]&Lvl[B]); for Type, disjointness is on the resolved type sets
'       (including the "any cell" token);
'   (b) a positive Cell[...] (no wildcard) that requires a cell coexisting with a Type condition that
'       forbids cells (positive Type with no cell in its set, or a negated Type[...] covering all cells).
' Any wildcard in a candidate contradiction => NO verdict (return False). Narrow by design (the mission's
' four cases), conservative (never flags a rule that could match). Used by the 13-3 coloured preview.
Public Function RuleHasNoEffect(ByVal sRule As String, ByRef segments() As String) As Boolean
    On Error GoTo ErrorHandler

    RuleHasNoEffect = False
    ReDim segments(0 To 0)
    segments(0) = ""

    Dim r As RuleInfo
    Dim sReason As String
    sReason = ParseOneRule(sRule, r)
    If Len(sReason) > 0 Then Exit Function       ' only meaningful on a valid rule
    If r.nCond < 2 Then Exit Function            ' need >= 2 conditions to contradict

    Dim i As Long, j As Long
    For i = 0 To r.nCond - 2
        For j = i + 1 To r.nCond - 1
            If PairContradicts(r.Conditions(i), r.Conditions(j)) Then
                ReDim segments(0 To 1)
                segments(0) = ConditionToCanonical(r.Conditions(i))
                segments(1) = ConditionToCanonical(r.Conditions(j))
                RuleHasNoEffect = True
                Exit Function
            End If
        Next j
    Next i
    Exit Function

ErrorHandler:
    ' Silent fail-closed (no log), matching the ElementMatchesAnyRule query-helper convention:
    ' a fault here only withholds an advisory verdict - it is not a fault the user can act on.
    RuleHasNoEffect = False
    ReDim segments(0 To 0)
    segments(0) = ""
End Function

'######################################################################################################################
'                                          PARSER (single source of truth)
'######################################################################################################################

' Parse ARES_Property_Rules into mRules once; cached until RefreshRules. Splits the raw value on the
' depth-0 ";" (a ";" is only ever a rule separator - it is forbidden inside [...]), then parses each rule
' via ParseOneRule; a rule that does not fit grammar v2 (including every v1 rule) is SKIPPED fail-closed
' (not counted, no attach) and logs nothing (a stored bad rule is not a fault).
Private Sub EnsureRulesParsed()
    On Error GoTo ErrorHandler

    If mbParsed Then Exit Sub
    mbParsed = True
    mnRuleCount = 0

    Dim sRaw As String
    sRaw = GetRulesRaw()
    If Len(Trim(sRaw)) = 0 Then Exit Sub

    Dim vRules() As String
    vRules = SplitTopLevel(sRaw, RULE_SEPARATOR)
    ReDim mRules(0 To UBound(vRules))

    Dim k As Long
    Dim r As RuleInfo
    Dim sReason As String
    For k = LBound(vRules) To UBound(vRules)
        If Len(Trim(vRules(k))) > 0 Then
            sReason = ParseOneRule(vRules(k), r)
            If Len(sReason) = 0 Then
                mRules(mnRuleCount) = r
                mnRuleCount = mnRuleCount + 1
            End If
        End If
    Next k
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging.EnsureRulesParsed"
    mnRuleCount = 0
End Sub

' The bracket-depth-aware core. Returns "" on success (fills r) or a targeted English reason. "=", "&",
' "@", "(", ")" are STRUCTURAL only at bracket depth 0; inside [...] they are literal name characters.
Private Function ParseOneRule(ByVal sInput As String, ByRef r As RuleInfo) As String
    On Error GoTo ErrorHandler

    ParseOneRule = ""

    ' Reset the target so a previous rule cannot leak in.
    r.IsGroup = False
    r.nCond = 0
    Erase r.Conditions
    Erase r.Props

    Dim s As String
    s = Trim(sInput)
    If Len(s) = 0 Then
        ParseOneRule = "empty rule"
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
                ParseOneRule = "malformed [...] (unbalanced ']')"
                Exit Function
            End If
        End If
    Next i
    If depth <> 0 Then
        ParseOneRule = "malformed [...] (unbalanced '[')"
        Exit Function
    End If

    ' First depth-0 "=".
    Dim eqPos As Long
    eqPos = FindTopLevelChar(s, SELECTOR_SEPARATOR)
    If eqPos = 0 Then
        ParseOneRule = "rule has no '=' (expected condition=prop|prop)"
        Exit Function
    End If

    Dim condSide As String, propSide As String
    condSide = Trim(Left(s, eqPos - 1))
    propSide = Trim(Mid(s, eqPos + 1))
    If Len(condSide) = 0 Then
        ParseOneRule = "empty condition side (before '=')"
        Exit Function
    End If
    If Len(propSide) = 0 Then
        ParseOneRule = "empty property side (after '=')"
        Exit Function
    End If

    ' Scan the condition side: collect + strip the depth-0 "@" (the group modifier, any position before
    ' "="), reject a depth-0 "(" / ")". Brackets keep their content literal.
    Dim condText As String, ch As String
    depth = 0
    condText = ""
    For i = 1 To Len(condSide)
        ch = Mid(condSide, i, 1)
        Select Case ch
            Case BRK_OPEN
                depth = depth + 1
                condText = condText & ch
            Case BRK_CLOSE
                depth = depth - 1
                condText = condText & ch
            Case CELL_GROUP_MARKER
                If depth = 0 Then
                    r.IsGroup = True
                Else
                    condText = condText & ch
                End If
            Case PAREN_OPEN, PAREN_CLOSE
                If depth = 0 Then
                    ParseOneRule = "'(' and ')' are reserved"
                    Exit Function
                Else
                    condText = condText & ch
                End If
            Case Else
                condText = condText & ch
        End Select
    Next i

    condText = Trim(condText)
    If Len(condText) = 0 Then
        ParseOneRule = "empty condition side (before '=')"
        Exit Function
    End If

    ' Split the condition text on the depth-0 "&" into segments; parse each.
    Dim segs() As String
    segs = SplitTopLevel(condText, COND_SEPARATOR)
    ReDim r.Conditions(0 To UBound(segs))

    Dim cnd As RuleCondition
    Dim seg As String
    Dim nc As Long
    nc = 0
    For i = LBound(segs) To UBound(segs)
        seg = Trim(segs(i))
        If Len(seg) = 0 Then
            ParseOneRule = "empty condition (a '&' with nothing beside it)"
            Exit Function
        End If
        Dim sCondReason As String
        sCondReason = ParseCondition(seg, cnd)
        If Len(sCondReason) > 0 Then
            ParseOneRule = sCondReason
            Exit Function
        End If
        r.Conditions(nc) = cnd
        nc = nc + 1
    Next i
    r.nCond = nc

    ' Property side: "|"-separated, everything literal ("@" literal). Reject a prop containing "=" or ";"
    ' (the "|"-instead-of-";" signature). Empty tokens are dropped; at least one non-empty prop required.
    Dim vRawProps As Variant, tok As String
    vRawProps = Split(propSide, ARESConstants.ARES_VAR_DELIMITER)
    For i = LBound(vRawProps) To UBound(vRawProps)
        tok = Trim(vRawProps(i))
        If Len(tok) > 0 Then
            If InStr(tok, SELECTOR_SEPARATOR) > 0 Then
                ParseOneRule = "property '" & tok & "' contains '=' - separate rules with ';' not '|'?"
                Exit Function
            End If
            If InStr(tok, RULE_SEPARATOR) > 0 Then
                ParseOneRule = "property '" & tok & "' contains ';'"
                Exit Function
            End If
        End If
    Next i

    r.Props = SplitTrim(propSide, ARESConstants.ARES_VAR_DELIMITER)
    If Len(r.Props(LBound(r.Props))) = 0 Then
        ParseOneRule = "empty property side (after '=')"
        Exit Function
    End If
    Exit Function

ErrorHandler:
    ParseOneRule = "invalid rule"
End Function

' Parse ONE condition segment "[!] Keyword[name|name|...]" into c. Returns "" on success or a reason.
Private Function ParseCondition(ByVal segInput As String, ByRef c As RuleCondition) As String
    On Error GoTo ErrorHandler

    ParseCondition = ""

    ' Reset so a previous condition cannot leak in.
    c.Negated = False
    c.MatchesAnyCell = False
    Erase c.Names
    Erase c.Types

    Dim seg As String
    seg = Trim(segInput)
    If Len(seg) = 0 Then
        ParseCondition = "empty condition"
        Exit Function
    End If

    If Left(seg, 1) = NEG_MARKER Then
        c.Negated = True
        seg = Trim(Mid(seg, 2))
        If Len(seg) = 0 Then
            ParseCondition = "empty condition after '!'"
            Exit Function
        End If
    End If

    Dim nOpen As Long, nClose As Long
    nOpen = InStr(seg, BRK_OPEN)
    If nOpen = 0 Then
        ParseCondition = "condition '" & seg & "' has no keyword (expected Lvl[..]/Cell[..]/Type[..])"
        Exit Function
    End If
    nClose = InStr(seg, BRK_CLOSE)
    If nClose <= nOpen Then
        ParseCondition = "malformed [...] in '" & seg & "'"
        Exit Function
    End If
    If nClose <> Len(seg) Then
        ParseCondition = "unexpected text after ']' in '" & seg & "'"
        Exit Function
    End If

    Dim sKw As String, body As String
    sKw = Trim(Left(seg, nOpen - 1))
    body = Mid(seg, nOpen + 1, nClose - nOpen - 1)

    Select Case UCase(sKw)
        Case "LVL"
            c.Keyword = rkLvl
        Case "CELL"
            c.Keyword = rkCell
        Case "TYPE"
            c.Keyword = rkType
        Case Else
            If Len(sKw) = 0 Then
                ParseCondition = "condition has no keyword (expected Lvl/Cell/Type)"
            Else
                ParseCondition = "unknown keyword '" & sKw & "' (expected Lvl/Cell/Type)"
            End If
            Exit Function
    End Select

    ' Forbidden characters inside [...] (";" and "[" - "]" cannot be here since nClose is the first "]").
    If InStr(body, RULE_SEPARATOR) > 0 Then
        ParseCondition = "';' not allowed inside [...]"
        Exit Function
    End If
    If InStr(body, BRK_OPEN) > 0 Then
        ParseCondition = "'[' not allowed inside [...]"
        Exit Function
    End If

    ' Split the body on "|" into trimmed, NON-EMPTY names (kept verbatim for Like).
    Dim vNames As Variant, nm As String, nCount As Long
    Dim namesOut() As String
    vNames = Split(body, ARESConstants.ARES_VAR_DELIMITER)
    ReDim namesOut(0 To UBound(vNames))
    nCount = 0
    Dim i As Long
    For i = LBound(vNames) To UBound(vNames)
        nm = Trim(vNames(i))
        If Len(nm) = 0 Then
            ParseCondition = "empty name in " & KeywordName(c.Keyword) & "[...]"
            Exit Function
        End If
        namesOut(nCount) = nm
        nCount = nCount + 1
    Next i
    ReDim Preserve namesOut(0 To nCount - 1)
    c.Names = namesOut

    ' Type resolution: each name resolves to an MsdElementType, or is the special "Cell" token (any cell).
    ' Wildcards are not meaningful for a type (an enum, not a name) -> rejected.
    If c.Keyword = rkType Then
        Dim typesOut() As Long
        Dim nt As Long
        ReDim typesOut(0 To nCount - 1)
        nt = 0
        For i = 0 To nCount - 1
            Dim bWild As Boolean
            bWild = False
            If InStr(c.Names(i), "*") > 0 Then bWild = True
            If InStr(c.Names(i), "?") > 0 Then bWild = True
            If bWild Then
                ParseCondition = "wildcards not allowed in Type[...]"
                Exit Function
            End If
            If UCase(c.Names(i)) = "CELL" Then
                c.MatchesAnyCell = True
            Else
                Dim t As Long
                t = MicroStationDefinition.StringToMsdElementType(c.Names(i))
                If t = 0 Then
                    ParseCondition = "unknown element type '" & c.Names(i) & "'"
                    Exit Function
                End If
                typesOut(nt) = t
                nt = nt + 1
            End If
        Next i
        If nt > 0 Then
            ReDim Preserve typesOut(0 To nt - 1)
            c.Types = typesOut
        End If
    End If
    Exit Function

ErrorHandler:
    ParseCondition = "invalid condition"
End Function

'######################################################################################################################
'                                          MATCHER
'######################################################################################################################

' Does the parsed rule match the element? AND over all conditions with strict negation. sLevel/bHasLevel
' are resolved once by the caller (guarded), so a level-less cell header still evaluates Cell/Type.
Private Function RuleMatches(ByRef r As RuleInfo, ByVal oElement As element, ByVal sLevel As String, ByVal bHasLevel As Boolean) As Boolean
    On Error GoTo ErrorHandler

    RuleMatches = False
    Dim i As Long
    For i = 0 To r.nCond - 1
        If Not ConditionMatches(r.Conditions(i), oElement, sLevel, bHasLevel) Then Exit Function
    Next i
    RuleMatches = True
    Exit Function

ErrorHandler:
    RuleMatches = False
End Function

' Evaluate ONE condition with strict negation. The positive result is computed with per-keyword guards
' (never And-chained across a possibly-raising read); a negated condition returns Not(positive). On a
' level-less element a positive Lvl is False (a negated !Lvl is True); on a non-cell a positive Cell is
' False (!Cell is True) - so "Type[Cell]&!Cell[A]" means exactly "a cell, but not the one named A".
Private Function ConditionMatches(ByRef c As RuleCondition, ByVal oElement As element, ByVal sLevel As String, ByVal bHasLevel As Boolean) As Boolean
    On Error GoTo ErrorHandler

    Dim bPos As Boolean
    bPos = False

    Select Case c.Keyword
        Case rkLvl
            If bHasLevel Then
                bPos = LikeAnyCI(sLevel, c.Names)
            End If
        Case rkCell
            If oElement.IsCellElement Then
                bPos = LikeAnyCI(oElement.AsCellElement.Name, c.Names)
            End If
        Case rkType
            If c.MatchesAnyCell Then
                If oElement.IsCellElement Then bPos = True
            End If
            If Not bPos Then
                If HasLongs(c.Types) Then
                    Dim ti As Long
                    For ti = LBound(c.Types) To UBound(c.Types)
                        If oElement.Type = c.Types(ti) Then
                            bPos = True
                            Exit For
                        End If
                    Next ti
                End If
            End If
    End Select

    If c.Negated Then
        ConditionMatches = Not bPos
    Else
        ConditionMatches = bPos
    End If
    Exit Function

ErrorHandler:
    ' Fail-closed: an unexpected fault counts as no match (never an errant attach).
    ConditionMatches = False
End Function

' Case-insensitive Like match of value against any of names. VBA Like metacharacters that could appear
' literally in a name are neutralised: only "#" can occur ("[" / "]" are forbidden inside a name), so
' escape "#" -> "[#]"; "*"/"?" stay wildcards. Case-insensitivity via UCase on both sides (the module has
' no Option Compare Text). Nested guards, no short-circuit.
Private Function LikeAnyCI(ByVal value As String, ByRef names() As String) As Boolean
    On Error GoTo ErrorHandler

    LikeAnyCI = False
    Dim uv As String
    uv = UCase(value)

    Dim i As Long
    For i = LBound(names) To UBound(names)
        If Len(names(i)) > 0 Then
            If uv Like UCase(EscapeLikePattern(names(i))) Then
                LikeAnyCI = True
                Exit Function
            End If
        End If
    Next i
    Exit Function

ErrorHandler:
    LikeAnyCI = False
End Function

' Escape the only Like metacharacter that can appear literally in a name: "#" -> "[#]". "*"/"?" are kept
' as wildcards; "[" / "]" cannot occur in a name (grammar-forbidden), so nothing else needs escaping.
Private Function EscapeLikePattern(ByVal name As String) As String
    EscapeLikePattern = Replace(name, "#", "[#]")
End Function

' Fan the rule's props out to each OTHER member of the element's graphic group (idempotent attach).
' Nothing without a real graphic group.
Private Sub AttachGroupMembers(ByVal oElement As element, ByRef r As RuleInfo)
    On Error GoTo ErrorHandler

    If oElement.GraphicGroup = ARES_DEFAULT_GRAPHIC_GROUP_ID Then Exit Sub

    Dim els() As element
    els = Link.GetLink(oElement)
    If Not HasElements(els) Then Exit Sub

    Dim j As Long, k As Long
    Dim s As element
    For j = LBound(els) To UBound(els)
        Set s = els(j)
        If Not s Is Nothing Then
            For k = LBound(r.Props) To UBound(r.Props)
                If Len(r.Props(k)) > 0 Then
                    CustomPropertyHandler.AttachItemToElement s, r.Props(k)
                End If
            Next k
        End If
    Next j
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging.AttachGroupMembers"
End Sub

'######################################################################################################################
'                                          CANONICALISATION
'######################################################################################################################

' Build the COMPACT canonical form of a parsed rule: [@] cond [&cond]* = prop[|prop]* with NO spaces
' around "&"/"=" and none inside [...], canonical keyword casing, names/props verbatim (already trimmed).
' NOTE: the canonical form matches the story's I/O matrix exactly (e.g. Type[Cell]&!Cell[A]=Repere) - a
' compact form, not the "single space around &" phrasing that appears in the Boundaries prose (the matrix
' is the contract).
Private Function RuleToCanonical(ByRef r As RuleInfo) As String
    Dim sOut As String
    Dim i As Long

    sOut = ""
    If r.IsGroup Then sOut = CELL_GROUP_MARKER

    For i = 0 To r.nCond - 1
        If i > 0 Then sOut = sOut & COND_SEPARATOR
        sOut = sOut & ConditionToCanonical(r.Conditions(i))
    Next i

    sOut = sOut & SELECTOR_SEPARATOR & Join(r.Props, ARESConstants.ARES_VAR_DELIMITER)
    RuleToCanonical = sOut
End Function

' Canonical text of one condition: [!] Keyword[name|name|...].
Private Function ConditionToCanonical(ByRef c As RuleCondition) As String
    Dim s As String
    s = ""
    If c.Negated Then s = NEG_MARKER
    s = s & KeywordName(c.Keyword)
    s = s & BRK_OPEN & Join(c.Names, ARESConstants.ARES_VAR_DELIMITER) & BRK_CLOSE
    ConditionToCanonical = s
End Function

' Canonical keyword casing.
Private Function KeywordName(ByVal kw As RuleKeyword) As String
    Select Case kw
        Case rkLvl
            KeywordName = "Lvl"
        Case rkCell
            KeywordName = "Cell"
        Case rkType
            KeywordName = "Type"
        Case Else
            KeywordName = ""
    End Select
End Function

'######################################################################################################################
'                                          CONTRADICTION DETECTOR
'######################################################################################################################

' True when two conditions can never be satisfied together (see RuleHasNoEffect for the covered cases).
Private Function PairContradicts(ByRef a As RuleCondition, ByRef b As RuleCondition) As Boolean
    On Error GoTo ErrorHandler

    PairContradicts = False

    ' (a) Same keyword, both positive, disjoint (no wildcard for Lvl/Cell; Type has none).
    If Not a.Negated Then
        If Not b.Negated Then
            If a.Keyword = b.Keyword Then
                Select Case a.Keyword
                    Case rkLvl, rkCell
                        If Not HasWildcard(a) Then
                            If Not HasWildcard(b) Then
                                If NamesDisjoint(a.Names, b.Names) Then
                                    PairContradicts = True
                                    Exit Function
                                End If
                            End If
                        End If
                    Case rkType
                        If TypeCondsDisjoint(a, b) Then
                            PairContradicts = True
                            Exit Function
                        End If
                End Select
            End If
        End If
    End If

    ' (b) Cell[...] (requires a cell) vs a Type condition that forbids cells - either order.
    If CellTypeContradict(a, b) Then
        PairContradicts = True
        Exit Function
    End If
    If CellTypeContradict(b, a) Then
        PairContradicts = True
        Exit Function
    End If
    Exit Function

ErrorHandler:
    PairContradicts = False
End Function

' c must be a positive Cell[...]; t a Type[...] that forbids cells: a positive Type whose resolved set
' contains no cell, or a negated Type[...] that covers all cells (!Type[Cell]). This contradiction is
' STRUCTURAL - "is the element a cell?" does not depend on the cell NAME - so there is NO wildcard guard
' here (unlike the same-keyword disjoint-list check, where a wildcard makes disjointness undecidable).
Private Function CellTypeContradict(ByRef c As RuleCondition, ByRef t As RuleCondition) As Boolean
    CellTypeContradict = False

    If c.Keyword <> rkCell Then Exit Function
    If c.Negated Then Exit Function
    If t.Keyword <> rkType Then Exit Function

    If Not t.Negated Then
        ' A positive Type with no cell in its match-set forbids the cell that Cell[...] requires.
        If Not t.MatchesAnyCell Then
            If Not HasCellType(t) Then
                CellTypeContradict = True
            End If
        End If
    Else
        ' A negated Type covering all cells (!Type[Cell]) forbids the cell Cell[...] requires.
        If t.MatchesAnyCell Then
            CellTypeContradict = True
        End If
    End If
End Function

' True when a condition carries a "*" or "?" in any of its names.
Private Function HasWildcard(ByRef c As RuleCondition) As Boolean
    On Error GoTo ErrorHandler

    HasWildcard = False
    Dim i As Long
    For i = LBound(c.Names) To UBound(c.Names)
        If InStr(c.Names(i), "*") > 0 Then
            HasWildcard = True
            Exit Function
        End If
        If InStr(c.Names(i), "?") > 0 Then
            HasWildcard = True
            Exit Function
        End If
    Next i
    Exit Function

ErrorHandler:
    HasWildcard = True                           ' fail-safe: treat as "wildcard present" -> no verdict
End Function

' True when name-lists a and b share no name (case-insensitive, trimmed) -> disjoint.
Private Function NamesDisjoint(ByRef a() As String, ByRef b() As String) As Boolean
    On Error GoTo ErrorHandler

    NamesDisjoint = True
    Dim i As Long, j As Long
    For i = LBound(a) To UBound(a)
        For j = LBound(b) To UBound(b)
            If StrComp(Trim(a(i)), Trim(b(j)), vbTextCompare) = 0 Then
                NamesDisjoint = False
                Exit Function
            End If
        Next j
    Next i
    Exit Function

ErrorHandler:
    NamesDisjoint = False                        ' fail-safe: assume they overlap -> no verdict
End Function

' True when two positive Type conditions can never match the same element (their match-sets are disjoint,
' accounting for the "any cell" token).
Private Function TypeCondsDisjoint(ByRef a As RuleCondition, ByRef b As RuleCondition) As Boolean
    On Error GoTo ErrorHandler

    TypeCondsDisjoint = False

    If a.MatchesAnyCell Then
        If b.MatchesAnyCell Then Exit Function       ' both any-cell -> overlap
        If HasCellType(b) Then Exit Function         ' a any-cell, b lists a cell type -> overlap
    End If
    If b.MatchesAnyCell Then
        If HasCellType(a) Then Exit Function         ' b any-cell, a lists a cell type -> overlap
    End If
    If TypesIntersect(a.Types, b.Types) Then Exit Function

    TypeCondsDisjoint = True
    Exit Function

ErrorHandler:
    TypeCondsDisjoint = False
End Function

' True when a resolved type list contains a cell type (CellHeader or SharedCell).
Private Function HasCellType(ByRef c As RuleCondition) As Boolean
    On Error GoTo ErrorHandler

    HasCellType = False
    If Not HasLongs(c.Types) Then Exit Function
    Dim i As Long
    For i = LBound(c.Types) To UBound(c.Types)
        If c.Types(i) = msdElementTypeCellHeader Then
            HasCellType = True
            Exit Function
        End If
        If c.Types(i) = msdElementTypeSharedCell Then
            HasCellType = True
            Exit Function
        End If
    Next i
    Exit Function

ErrorHandler:
    HasCellType = False
End Function

' True when two resolved type lists share a value.
Private Function TypesIntersect(ByRef a() As Long, ByRef b() As Long) As Boolean
    On Error GoTo ErrorHandler

    TypesIntersect = False
    If Not HasLongs(a) Then Exit Function
    If Not HasLongs(b) Then Exit Function
    Dim i As Long, j As Long
    For i = LBound(a) To UBound(a)
        For j = LBound(b) To UBound(b)
            If a(i) = b(j) Then
                TypesIntersect = True
                Exit Function
            End If
        Next j
    Next i
    Exit Function

ErrorHandler:
    TypesIntersect = False
End Function

'######################################################################################################################
'                                          LOW-LEVEL HELPERS
'######################################################################################################################

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

' Position (1-based) of the first occurrence of the single character ch at bracket depth 0, or 0 if none.
Private Function FindTopLevelChar(ByVal s As String, ByVal ch As String) As Long
    Dim depth As Long, i As Long, c As String
    depth = 0
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c = BRK_OPEN Then
            depth = depth + 1
        ElseIf c = BRK_CLOSE Then
            depth = depth - 1
        ElseIf c = ch Then
            If depth = 0 Then
                FindTopLevelChar = i
                Exit Function
            End If
        End If
    Next i
    FindTopLevelChar = 0
End Function

' Split s on the single character ch at bracket depth 0 (a ch inside [...] is literal). Returns a 0-based
' array of the raw (untrimmed) segments, including empties.
Private Function SplitTopLevel(ByVal s As String, ByVal ch As String) As String()
    Dim out() As String
    Dim n As Long, depth As Long, i As Long, c As String, seg As String
    ReDim out(0 To 0)
    n = 0
    depth = 0
    seg = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c = BRK_OPEN Then
            depth = depth + 1
            seg = seg & c
        ElseIf c = BRK_CLOSE Then
            depth = depth - 1
            seg = seg & c
        ElseIf c = ch Then
            If depth = 0 Then
                If n > UBound(out) Then ReDim Preserve out(0 To n)
                out(n) = seg
                n = n + 1
                seg = ""
            Else
                seg = seg & c
            End If
        Else
            seg = seg & c
        End If
    Next i
    If n > UBound(out) Then ReDim Preserve out(0 To n)
    out(n) = seg
    SplitTopLevel = out
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

' Safe "String array has at least one element" check.
Private Function HasElements(ByRef arr() As element) As Boolean
    On Error Resume Next
    HasElements = False
    If UBound(arr) <> -1 Then HasElements = True
    On Error GoTo 0
End Function

' Safe "Long array is allocated and non-empty" check (mirrors HasElements for Types()).
Private Function HasLongs(ByRef arr() As Long) As Boolean
    On Error Resume Next
    HasLongs = False
    If UBound(arr) <> -1 Then HasLongs = True
    On Error GoTo 0
End Function
