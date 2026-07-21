' Module: PropertyTagging
' Description: Auto-attaches ARES custom properties to elements as they are created / modified,
'              driven by configurable rules (ARES_Property_Rules). Each rule maps a selector
'              (level, plus an optional element type) to a list of property names. Attach-only: the
'              property appears (empty) on the element for the user to fill / pick from its native
'              value-list dropdown.
'
'              Called from ElementChangeHandler.ProcessElement (deferred, on idle) when
'              ARES_Auto_Properties is True. Rules are parsed once and cached - call RefreshRules
'              after changing ARES_Property_Rules at runtime.
'
'              Rule format (ARES_Property_Rules):  "level[:type]=prop|prop ; level[:type]=prop ; ..."
'                - level : MicroStation level name (required; case-insensitive)
'                - type  : optional element type name (StringToMsdElementType); absent = any type
'                - props : | -delimited property names (must exist in the "ARES" DGNLib)
'              Example:  WALLS=Commune|Coupe_Type ; DOORS:Cell=Commune
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (global ARESConfig), CustomPropertyHandler, MicroStationDefinition, ErrorHandlerClass (global ErrorHandler)

Option Explicit

Private Const RULE_SEPARATOR As String = ";"
Private Const SELECTOR_SEPARATOR As String = "="
Private Const TYPE_SEPARATOR As String = ":"

' One parsed rule: a level (required) + optional element type -> property names to attach.
Private Type RuleInfo
    Level As String
    HasType As Boolean
    ElType As MsdElementType
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

    For i = 0 To mnRuleCount - 1
        If RuleMatches(mRules(i), oElement, sLevel) Then
            ElementMatchesAnyRule = True
            Exit Function
        End If
    Next i
    Exit Function

ErrorHandler:
    ElementMatchesAnyRule = False
End Function

' Attach the configured properties to the element for every rule it matches. Idempotent
' (CustomPropertyHandler.AttachItemToElement only attaches when not already present).
Public Sub ApplyPropertyRules(ByVal oElement As element)
    On Error GoTo ErrorHandler

    If oElement Is Nothing Then Exit Sub

    EnsureRulesParsed
    If mnRuleCount = 0 Then Exit Sub

    Dim sLevel As String
    Dim i As Long, j As Long
    ' Level-based rules: an element with no level matches nothing. Guard with nested Ifs (never And -
    ' no short-circuit in VBA): Level GET raises on a non-graphical element and can be Nothing (e.g. a
    ' cell header), which would make .Name raise Error 91. Silent skip, no log.
    If Not oElement.IsGraphical Then Exit Sub
    If oElement.Level Is Nothing Then Exit Sub
    sLevel = oElement.Level.Name

    For i = 0 To mnRuleCount - 1
        If RuleMatches(mRules(i), oElement, sLevel) Then
            For j = LBound(mRules(i).Props) To UBound(mRules(i).Props)
                If Len(mRules(i).Props(j)) > 0 Then
                    CustomPropertyHandler.AttachItemToElement oElement, mRules(i).Props(j)
                End If
            Next j
        End If
    Next i
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging.ApplyPropertyRules"
End Sub

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

                ' selector = level[:type]
                sType = ""
                nColon = InStr(sSelector, TYPE_SEPARATOR)
                If nColon > 0 Then
                    sType = Trim(Mid(sSelector, nColon + 1))
                    sSelector = Trim(Left(sSelector, nColon - 1))
                End If

                If Len(sSelector) > 0 And Len(sPropsRaw) > 0 Then
                    r.Level = sSelector
                    r.HasType = (Len(sType) > 0)
                    If r.HasType Then r.ElType = MicroStationDefinition.StringToMsdElementType(sType)
                    r.Props = SplitTrim(sPropsRaw, ARESConstants.ARES_VAR_DELIMITER)
                    mRules(mnRuleCount) = r
                    mnRuleCount = mnRuleCount + 1
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
