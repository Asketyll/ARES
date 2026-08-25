' Module: PropertyTagging
' Description: The SOLE attach/detach engine for ARES custom properties. Auto-attaches properties to
'              elements as they are created/modified, driven by configurable rules (ARES_Property_Rules).
'              Attach-only on match (empty value, filled by the user or by Calculation). Called from
'              ElementChangeHandler.ProcessElement (deferred, on idle) when ARES_Auto_Properties is True;
'              rules are parsed once and cached - call RefreshRules after changing the variable at runtime.
'              Grammar v2, the "@" group-membership convergence mechanism, the RuleGrammar delegation, and
'              the attach/detach choke point (incl. render metadata): see
'              _bmad/docs/property-tagging-grammar.md.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (global ARESConfig), CustomPropertyHandler, Link, RuleGrammar, ErrorHandlerClass (global ErrorHandler), CallStackClass (global CallStack)

Option Explicit

Private Const RULE_SEPARATOR As String = ";"
Private Const SELECTOR_SEPARATOR As String = "="
Private Const COND_SEPARATOR As String = "&"
Private Const CELL_GROUP_MARKER As String = "@"
Private Const BRK_OPEN As String = "["
Private Const BRK_CLOSE As String = "]"
Private Const PAREN_OPEN As String = "("
Private Const PAREN_CLOSE As String = ")"
' The name separator inside [...] and between property names is ARESConstants.ARES_VAR_DELIMITER ("|").

' One parsed rule: [@] conditions (AND) = props. IsGroup is the "@" modifier (attach to the OTHER
' graphic-group members). nCond bounds the meaningful entries of Conditions() (RuleGrammar.RuleCondition).
Private Type RuleInfo
    IsGroup As Boolean
    Conditions() As RuleGrammar.RuleCondition
    nCond As Long
    Props() As String
End Type

Private mRules() As RuleInfo
Private mnRuleCount As Long
Private mbParsed As Boolean
' True as soon as ONE parsed rule carries the "@" modifier. Cost guard: without it the group work (the
' shared Link.GetLink fetch, the push fan-out and the pull pass) is never entered, so a config with no "@"
' rule pays nothing. Recomputed by EnsureRulesParsed, which RefreshRules forces to re-run.
Private mbHasGroupRules As Boolean

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

' Attach the configured properties for every rule the element drives, in TWO symmetric directions (PUSH /
' PULL). Full mechanism: see "ApplyPropertyRules - PUSH/PULL mechanism" in property-tagging-grammar.md.
Public Sub ApplyPropertyRules(ByVal oElement As element)
    On Error GoTo ErrorHandler

    Dim bStackPushed As Boolean
    If oElement Is Nothing Then Exit Sub

    EnsureRulesParsed
    If mnRuleCount = 0 Then Exit Sub

    CallStack.Push "PropertyTagging.ApplyPropertyRules", oElement
    bStackPushed = True

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

    ' ONE member fetch per processed element, shared by the push and the pull. NESTED Ifs, never "And" (VBA
    ' does not short-circuit, .GraphicGroup raises on a non-graphical element). See property-tagging-grammar.md.
    Dim els() As element
    Dim bHaveEls As Boolean
    bHaveEls = False
    If mbHasGroupRules Then
        If oElement.IsGraphical Then
            If oElement.GraphicGroup <> ARES_DEFAULT_GRAPHIC_GROUP_ID Then
                els = Link.GetLink(oElement)
                bHaveEls = HasElements(els)
            End If
        End If
    End If

    Dim i As Long, j As Long
    For i = 0 To mnRuleCount - 1
        If RuleMatches(mRules(i), oElement, sLevel, bHasLevel) Then
            If mRules(i).IsGroup Then
                AttachGroupMembers oElement, mRules(i), els, bHaveEls
            Else
                For j = LBound(mRules(i).Props) To UBound(mRules(i).Props)
                    If Len(mRules(i).Props(j)) > 0 Then
                        CustomPropertyHandler.AttachItemToElement oElement, mRules(i).Props(j)
                    End If
                Next j
            End If
        End If
    Next i

    ' PULL pass - runs after the push, on the SAME members array.
    If bHaveEls Then PullGroupProperties oElement, els
    CallStack.Pop
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging.ApplyPropertyRules"
    If bStackPushed Then CallStack.Pop
End Sub

' Public detach service: remove a single property P from El (ARES_Calc_Detach_Empty). Thin wrapper over
' CustomPropertyHandler.RemoveItemFromElement, idempotent. See property-tagging-grammar.md.
Public Sub DetachRuleProperty(ByVal El As element, ByVal P As String)
    On Error GoTo ErrorHandler

    If El Is Nothing Then Exit Sub
    If Len(Trim(P)) = 0 Then Exit Sub

    CustomPropertyHandler.RemoveItemFromElement El, P
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging.DetachRuleProperty"
End Sub

' Public attach service for the RENDERER's internal metadata (epic 15). PropertyRendering never attaches
' anything itself: the attach choke point stays unique, exactly as the value engine's detach does above.
' The LibraryName is passed EXPLICITLY - CustomPropertyHandler defaults it to the user-facing "ARES"
' library, which does not hold this ItemType. Idempotent (HasItems-guarded downstream).
Public Function AttachRenderMetadata(ByVal El As element) As Boolean
    On Error GoTo ErrorHandler

    AttachRenderMetadata = False
    If El Is Nothing Then Exit Function

    AttachRenderMetadata = CustomPropertyHandler.AttachItemToElement(El, ARES_ITEM_RENDER, ARES_NAME_LIBRARY_SYS)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging.AttachRenderMetadata"
    AttachRenderMetadata = False
End Function

' Public detach service for the renderer's internal metadata - the mirror of AttachRenderMetadata.
' Same explicit-library rule. Idempotent (RemoveItemFromElement is HasItems-guarded).
Public Sub DetachRenderMetadata(ByVal El As element)
    On Error GoTo ErrorHandler

    If El Is Nothing Then Exit Sub

    CustomPropertyHandler.RemoveItemFromElement El, ARES_ITEM_RENDER, ARES_NAME_LIBRARY_SYS
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging.DetachRenderMetadata"
End Sub

' Read-only validate-AND-normalise for ONE rule, the seam the editor writes through. Calls the SAME
' ParseOneRule the runtime parser uses (no drift). Syntactic only - no DGNLib membership check. See
' "Options-form services" in property-tagging-grammar.md.
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

' Read-only contradiction (dead-rule) detector on a SYNTACTICALLY VALID rule; delegates the reasoning to
' RuleGrammar.ConditionsHaveContradiction. Used by the coloured preview. See property-tagging-grammar.md.
Public Function RuleHasNoEffect(ByVal sRule As String, ByRef segments() As String) As Boolean
    On Error GoTo ErrorHandler

    RuleHasNoEffect = False
    ReDim segments(0 To 0)
    segments(0) = ""

    Dim r As RuleInfo
    Dim sReason As String
    sReason = ParseOneRule(sRule, r)
    If Len(sReason) > 0 Then Exit Function       ' only meaningful on a valid rule

    RuleHasNoEffect = RuleGrammar.ConditionsHaveContradiction(r.Conditions, r.nCond, segments)
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
    mbHasGroupRules = False

    Dim sRaw As String
    sRaw = GetRulesRaw()
    If Len(Trim(sRaw)) = 0 Then Exit Sub

    Dim vRules() As String
    vRules = RuleGrammar.SplitTopLevel(sRaw, RULE_SEPARATOR)
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
                If r.IsGroup Then mbHasGroupRules = True
            End If
        End If
    Next k
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging.EnsureRulesParsed"
    mnRuleCount = 0
    mbHasGroupRules = False
End Sub

' The rule shell + bracket-depth-aware "@"/"="/"&" split. Returns "" on success (fills r) or a targeted
' English reason. "=", "&", "@", "(", ")" are STRUCTURAL only at bracket depth 0; inside [...] they are
' literal name characters. Each condition segment is parsed by RuleGrammar.ParseCondition.
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
    eqPos = RuleGrammar.FindTopLevelChar(s, SELECTOR_SEPARATOR)
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

    ' Split the condition text on the depth-0 "&" into segments; parse each via RuleGrammar.
    Dim segs() As String
    segs = RuleGrammar.SplitTopLevel(condText, COND_SEPARATOR)
    ReDim r.Conditions(0 To UBound(segs))

    Dim cnd As RuleGrammar.RuleCondition
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
        sCondReason = RuleGrammar.ParseCondition(seg, cnd)
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

    r.Props = RuleGrammar.SplitTrim(propSide, ARESConstants.ARES_VAR_DELIMITER)
    If Len(r.Props(LBound(r.Props))) = 0 Then
        ParseOneRule = "empty property side (after '=')"
        Exit Function
    End If
    Exit Function

ErrorHandler:
    ParseOneRule = "invalid rule"
End Function

'######################################################################################################################
'                                          MATCHER
'######################################################################################################################

' Does the parsed rule match the element? AND over all conditions with strict negation (each condition via
' RuleGrammar.ConditionMatches). sLevel/bHasLevel are resolved once by the caller (guarded), so a
' level-less cell header still evaluates Cell/Type.
Private Function RuleMatches(ByRef r As RuleInfo, ByVal oElement As element, ByVal sLevel As String, ByVal bHasLevel As Boolean) As Boolean
    On Error GoTo ErrorHandler

    RuleMatches = False
    Dim i As Long
    For i = 0 To r.nCond - 1
        If Not RuleGrammar.ConditionMatches(r.Conditions(i), oElement, sLevel, bHasLevel) Then Exit Function
    Next i
    RuleMatches = True
    Exit Function

ErrorHandler:
    RuleMatches = False
End Function

' PUSH - fan the rule's props out to each OTHER member of the element's graphic group. See
' "AttachGroupMembers (PUSH)" in property-tagging-grammar.md (the bHaveEls / els() unallocated guard).
Private Sub AttachGroupMembers(ByVal oElement As element, ByRef r As RuleInfo, ByRef els() As element, ByVal bHaveEls As Boolean)
    On Error GoTo ErrorHandler

    If Not bHaveEls Then Exit Sub

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

' PULL - the exact mirror of AttachGroupMembers, membership-convergent. Full mechanism (MEMBERS-outer/
' RULES-inner shape, delivered() early exit): see "PullGroupProperties (PULL)" in property-tagging-grammar.md.
Private Sub PullGroupProperties(ByVal oElement As element, ByRef els() As element)
    On Error GoTo ErrorHandler

    If mnRuleCount = 0 Then Exit Sub

    ' Only "@" rules can deliver; nRemaining drives the early exit.
    Dim i As Long
    Dim nRemaining As Long
    nRemaining = 0
    For i = 0 To mnRuleCount - 1
        If mRules(i).IsGroup Then nRemaining = nRemaining + 1
    Next i
    If nRemaining = 0 Then Exit Sub

    Dim delivered() As Boolean
    ReDim delivered(0 To mnRuleCount - 1)

    Dim j As Long, k As Long
    Dim el As element
    Dim sLevel As String
    Dim bHasLevel As Boolean

    For j = LBound(els) To UBound(els)
        Set el = els(j)
        If Not el Is Nothing Then
            sLevel = ""
            bHasLevel = False
            If el.IsGraphical Then
                If Not el.Level Is Nothing Then
                    sLevel = el.Level.Name
                    bHasLevel = True
                End If
            End If

            For i = 0 To mnRuleCount - 1
                If mRules(i).IsGroup Then
                    If Not delivered(i) Then
                        If RuleMatches(mRules(i), el, sLevel, bHasLevel) Then
                            For k = LBound(mRules(i).Props) To UBound(mRules(i).Props)
                                If Len(mRules(i).Props(k)) > 0 Then
                                    CustomPropertyHandler.AttachItemToElement oElement, mRules(i).Props(k)
                                End If
                            Next k
                            delivered(i) = True
                            nRemaining = nRemaining - 1
                        End If
                    End If
                End If
            Next i

            If nRemaining = 0 Then Exit For
        End If
    Next j
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging.PullGroupProperties"
End Sub

'######################################################################################################################
'                                          CANONICALISATION
'######################################################################################################################

' Build the COMPACT canonical form of a parsed rule: [@] cond [&cond]* = prop[|prop]* with NO spaces
' around "&"/"=" and none inside [...], canonical keyword casing, names/props verbatim (already trimmed).
' Each condition's canonical text comes from RuleGrammar.ConditionToCanonical.
Private Function RuleToCanonical(ByRef r As RuleInfo) As String
    Dim sOut As String
    Dim i As Long

    sOut = ""
    If r.IsGroup Then sOut = CELL_GROUP_MARKER

    For i = 0 To r.nCond - 1
        If i > 0 Then sOut = sOut & COND_SEPARATOR
        sOut = sOut & RuleGrammar.ConditionToCanonical(r.Conditions(i))
    Next i

    sOut = sOut & SELECTOR_SEPARATOR & Join(r.Props, ARESConstants.ARES_VAR_DELIMITER)
    RuleToCanonical = sOut
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

' Safe "element array has at least one element" check.
Private Function HasElements(ByRef arr() As element) As Boolean
    On Error Resume Next
    HasElements = False
    If UBound(arr) <> -1 Then HasElements = True
    On Error GoTo 0
End Function
