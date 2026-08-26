' Module: PropertyCalculation_TriggerPush
' Description: Trigger-cell / trigger-level push mechanics - pushing a trigger element's attributes to the
'              OTHER graphic-group members MicroStation did not re-queue. No module-scope state; reaches
'              the parsed rule cache only through Core's Public RuleCount()/GetRule() accessors, never the
'              raw array. Full grammar, engine passes and coupling doctrine: see
'              _bmad/docs/calc-rules-grammar.md.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: CustomPropertyHandler, Link, ErrorHandlerClass (global ErrorHandler),
'               PropertyCalculation (Core - RuleCount/GetRule/FindCalcRuleForProperty/ApplyValueToSibling/
'               ReportNoTarget/ReportMultipleTriggers/ReportMultipleLvlTriggers/HasElements),
'               PropertyCalculation_SourceEval (ReadCellSourceValue/ReadLvlSourceValue/MatchesAnyPattern)

Option Explicit

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

' Trigger-cell pass: pushes oCell's attributes to OTHER group members whose first-matching rule for their
' target P is fed by oCell's name (AC3 first-match guard). Each SourceKind is read at most once per call
' (cached by enum ordinal). No carrying member -> CalculationNoTarget; two competing cells -> Multiple.
' Iterates the rule cache through Core.RuleCount()/Core.GetRule() (never the raw array) - a real code
' change from the original file, not pure code motion, per the split plan §1 Module C.
Public Sub PushCellDerivedValuesToMembers(ByVal oCell As element)
    On Error GoTo ErrorHandler

    Dim members() As element
    members = Link.GetLink(oCell)                 ' OTHER members (self handled by the bearing pass)
    If Not PropertyCalculation.HasElements(members) Then Exit Sub

    Dim sName As String
    sName = oCell.AsCellElement.Name

    ' Lazy per-SourceKind cache, indexed by the CalcSource enum ordinal (csGroupLength is the highest).
    Dim cacheVal(0 To csGroupLength) As String
    Dim cacheReady(0 To csGroupLength) As Boolean

    Dim i As Long, ri As Long, kIdx As Long
    Dim m As element
    Dim P As String
    Dim rule As CalcRuleInfo
    Dim nCarried As Long
    nCarried = 0
    For i = LBound(members) To UBound(members)
        Set m = members(i)
        If Not m Is Nothing Then
            For ri = 0 To PropertyCalculation.RuleCount() - 1
                rule = PropertyCalculation.GetRule(ri)
                If IsPushableCellSourceKind(rule.SourceKind) Then
                    If PropertyCalculation_SourceEval.MatchesAnyPattern(sName, rule.SourceArg) Then
                        P = rule.TargetProp
                        If CustomPropertyHandler.IsItemAttachedToElement(m, P) Then
                            nCarried = nCarried + 1
                            ' First-match guard (AC3): push only where THIS rule governs m's P.
                            If PropertyCalculation.FindCalcRuleForProperty(P, m) = ri Then
                                kIdx = CLng(rule.SourceKind)
                                If Not cacheReady(kIdx) Then
                                    cacheVal(kIdx) = PropertyCalculation_SourceEval.ReadCellSourceValue(oCell, rule.SourceKind)
                                    cacheReady(kIdx) = True
                                End If
                                PropertyCalculation.ApplyValueToSibling m, P, cacheVal(kIdx)
                            End If
                        End If
                    End If
                End If
            Next ri
        End If
    Next i

    ' Discoverability: siblings match a rule but NONE carry its target -> attach never happened.
    If nCarried = 0 Then PropertyCalculation.ReportNoTarget

    ' Multi-trigger: another cell in the group also feeds one of the pushed targets (last-processed wins).
    If GroupHasCompetingTrigger(oCell) Then PropertyCalculation.ReportMultipleTriggers
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_TriggerPush.PushCellDerivedValuesToMembers"
End Sub

' True when oCell's graphic group holds at least one OTHER cell that also matches some pushable Cell*
' source's pattern (i.e. two label/anchor cells could feed the same target) - the multi-trigger condition.
' Read-only.
Private Function GroupHasCompetingTrigger(ByVal oCell As element) As Boolean
    On Error GoTo ErrorHandler

    GroupHasCompetingTrigger = False

    Dim members() As element
    members = Link.GetLink(oCell)                 ' OTHER members only
    If Not PropertyCalculation.HasElements(members) Then Exit Function

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
' that only feeds a CellId rule must not be treated as a trigger. Public: called from Core's IsTriggerCell.
' Iterates the rule cache through Core.RuleCount()/Core.GetRule() (never the raw array) - a real code
' change from the original file, not pure code motion, per the split plan §1 Module C.
Public Function AnyPushableSourcePatternMatches(ByVal sName As String) As Boolean
    On Error GoTo ErrorHandler

    AnyPushableSourcePatternMatches = False
    Dim i As Long
    Dim rule As CalcRuleInfo
    For i = 0 To PropertyCalculation.RuleCount() - 1
        rule = PropertyCalculation.GetRule(i)
        If IsPushableCellSourceKind(rule.SourceKind) Then
            If PropertyCalculation_SourceEval.MatchesAnyPattern(sName, rule.SourceArg) Then
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
'                                          ENGINE - TRIGGER-LEVEL PASS
'######################################################################################################################

' Parallel trigger/push pass for Lvl* sources, deliberately DUPLICATED rather than folded into Cell*: the
' trigger predicate differs structurally (Cell* = a cell whose NAME matches; Lvl* = any element whose
' LEVEL matches). KNOWN LIMITATION: a competing Cell*-trigger and Lvl*-trigger for the same target are not
' cross-detected (each family only scans its own kind) - the write itself stays safe (first-match-guarded).

' True for the Lvl* source kinds. Unlike Cell*, there is no LvlId to exclude - all three are pushable.
Private Function IsPushableLvlSourceKind(ByVal kind As CalcSource) As Boolean
    IsPushableLvlSourceKind = (kind = csLvlColor Or kind = csLvlStyle Or kind = csLvlWeight)
End Function

' Trigger-level pass: oTriggerEl is a trigger whose color/style/weight may have changed while its group
' members were NOT re-queued. Mirrors PushCellDerivedValuesToMembers exactly, substituting the level-name
' match for the cell-name match and ReadLvlSourceValue for ReadCellSourceValue. Iterates the rule cache
' through Core.RuleCount()/Core.GetRule() (never the raw array) - a real code change, per the split plan.
Public Sub PushLvlDerivedValuesToMembers(ByVal oTriggerEl As element)
    On Error GoTo ErrorHandler

    Dim members() As element
    members = Link.GetLink(oTriggerEl)             ' OTHER members (self handled by the bearing pass)
    If Not PropertyCalculation.HasElements(members) Then Exit Sub

    Dim sLevelName As String
    sLevelName = oTriggerEl.Level.Name

    ' Lazy per-SourceKind cache, indexed by the CalcSource enum ordinal (csGroupLength is the highest).
    Dim cacheVal(0 To csGroupLength) As String
    Dim cacheReady(0 To csGroupLength) As Boolean

    Dim i As Long, ri As Long, kIdx As Long
    Dim m As element
    Dim P As String
    Dim rule As CalcRuleInfo
    Dim nCarried As Long
    nCarried = 0
    For i = LBound(members) To UBound(members)
        Set m = members(i)
        If Not m Is Nothing Then
            For ri = 0 To PropertyCalculation.RuleCount() - 1
                rule = PropertyCalculation.GetRule(ri)
                If IsPushableLvlSourceKind(rule.SourceKind) Then
                    If PropertyCalculation_SourceEval.MatchesAnyPattern(sLevelName, rule.SourceArg) Then
                        P = rule.TargetProp
                        If CustomPropertyHandler.IsItemAttachedToElement(m, P) Then
                            nCarried = nCarried + 1
                            ' First-match guard (AC3): push only where THIS rule governs m's P.
                            If PropertyCalculation.FindCalcRuleForProperty(P, m) = ri Then
                                kIdx = CLng(rule.SourceKind)
                                If Not cacheReady(kIdx) Then
                                    cacheVal(kIdx) = PropertyCalculation_SourceEval.ReadLvlSourceValue(oTriggerEl, rule.SourceKind)
                                    cacheReady(kIdx) = True
                                End If
                                PropertyCalculation.ApplyValueToSibling m, P, cacheVal(kIdx)
                            End If
                        End If
                    End If
                End If
            Next ri
        End If
    Next i

    ' Discoverability: siblings match a rule but NONE carry its target -> attach never happened.
    If nCarried = 0 Then PropertyCalculation.ReportNoTarget

    ' Multi-trigger: another LEVEL-matching element in the group also feeds one of the pushed targets
    ' (last-processed wins). Own DISTINCT status (ReportMultipleLvlTriggers, not ReportMultipleTriggers):
    ' this collision may be pure Lvl-vs-Lvl, involving no cell at all - the message must not claim "cells".
    ' See the KNOWN LIMITATION note above the section header - cross-family (Cell-vs-Lvl) competition is
    ' still not detected here (a distinct, accepted gap, not this wording fix).
    If GroupHasCompetingLvlTrigger(oTriggerEl) Then PropertyCalculation.ReportMultipleLvlTriggers
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_TriggerPush.PushLvlDerivedValuesToMembers"
End Sub

' True when oTriggerEl's graphic group holds at least one OTHER element whose Level also matches some
' pushable Lvl* source's pattern - the multi-trigger condition, mirrors GroupHasCompetingTrigger. NO
' element-type restriction (mirrors IsTriggerLevel).
Private Function GroupHasCompetingLvlTrigger(ByVal oTriggerEl As element) As Boolean
    On Error GoTo ErrorHandler

    GroupHasCompetingLvlTrigger = False

    Dim members() As element
    members = Link.GetLink(oTriggerEl)             ' OTHER members only
    If Not PropertyCalculation.HasElements(members) Then Exit Function

    Dim i As Long
    For i = LBound(members) To UBound(members)
        If Not members(i) Is Nothing Then
            If members(i).IsGraphical Then
                If Not members(i).Level Is Nothing Then
                    If AnyPushableLvlSourcePatternMatches(members(i).Level.Name) Then
                        GroupHasCompetingLvlTrigger = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
    Exit Function

ErrorHandler:
    GroupHasCompetingLvlTrigger = False
End Function

' True when sLevelName matches a pushable Lvl* source's pattern of at least one calc rule (assumes the
' cache is parsed). Mirrors AnyPushableSourcePatternMatches. Public: called from Core's IsTriggerLevel.
' Iterates the rule cache through Core.RuleCount()/Core.GetRule() (never the raw array) - a real code
' change from the original file, not pure code motion, per the split plan §1 Module C.
Public Function AnyPushableLvlSourcePatternMatches(ByVal sLevelName As String) As Boolean
    On Error GoTo ErrorHandler

    AnyPushableLvlSourcePatternMatches = False
    Dim i As Long
    Dim rule As CalcRuleInfo
    For i = 0 To PropertyCalculation.RuleCount() - 1
        rule = PropertyCalculation.GetRule(i)
        If IsPushableLvlSourceKind(rule.SourceKind) Then
            If PropertyCalculation_SourceEval.MatchesAnyPattern(sLevelName, rule.SourceArg) Then
                AnyPushableLvlSourcePatternMatches = True
                Exit Function
            End If
        End If
    Next i
    Exit Function

ErrorHandler:
    AnyPushableLvlSourcePatternMatches = False
End Function
