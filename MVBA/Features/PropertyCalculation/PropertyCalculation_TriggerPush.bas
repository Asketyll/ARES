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
    ' Target properties oCell actually pushed a value to (post AC3 guard) - the only props for which
    ' another trigger cell in the group can be a REAL competitor. See GroupHasCompetingTrigger.
    Dim pushedTargets As New Collection
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
                                Dim sPushVal As String
                                If Len(rule.SourceSystem) > 0 Then
                                    ' A geo-output CellCoord rule (any requested system, or in principle any future rule carrying
                                    ' its own extra params) - the kIdx-only cache assumes "same SourceKind
                                    ' -> same value", which no longer holds once CellCoord rules can differ
                                    ' by SourceSystem/SourceGeoDecimals (see the WGS84 plan's §3.1). Bypass
                                    ' the shared cache entirely for this case and re-read every time;
                                    ' correctness over the cache's micro-optimisation, and this only runs on
                                    ' an actual trigger-cell edit, never a bulk hot path.
                                    sPushVal = PropertyCalculation_SourceEval.ReadCellSourceValue(oCell, rule.SourceKind, rule.SourceSystem, rule.SourceGeoDecimals)
                                Else
                                    If Not cacheReady(kIdx) Then
                                        cacheVal(kIdx) = PropertyCalculation_SourceEval.ReadCellSourceValue(oCell, rule.SourceKind)
                                        cacheReady(kIdx) = True
                                    End If
                                    sPushVal = cacheVal(kIdx)
                                End If
                                PropertyCalculation.ApplyValueToSibling m, P, sPushVal
                                If Not CollectionContainsString(pushedTargets, P) Then pushedTargets.Add P
                            End If
                        End If
                    End If
                End If
            Next ri
        End If
    Next i

    ' Discoverability: siblings match a rule but NONE carry its target -> attach never happened.
    If nCarried = 0 Then PropertyCalculation.ReportNoTarget

    ' Multi-trigger: another cell in the group also feeds one of oCell's OWN pushed targets (last-processed
    ' wins). Restricted to pushedTargets - two cells feeding two DIFFERENT properties are not competitors.
    If GroupHasCompetingTrigger(oCell, pushedTargets) Then PropertyCalculation.ReportMultipleTriggers
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_TriggerPush.PushCellDerivedValuesToMembers"
End Sub

' True when oCell's graphic group holds at least one OTHER cell that ALSO matches a pushable Cell* source
' feeding one of pushedTargets - the properties oCell itself just pushed a value to. Restricted to
' pushedTargets (not "any pushable rule") so that two cells feeding two DIFFERENT properties are not
' reported as competing - only a real collision on the SAME target property is. Read-only.
Private Function GroupHasCompetingTrigger(ByVal oCell As element, ByVal pushedTargets As Collection) As Boolean
    On Error GoTo ErrorHandler

    GroupHasCompetingTrigger = False
    If pushedTargets.Count = 0 Then Exit Function

    Dim members() As element
    members = Link.GetLink(oCell)                 ' OTHER members only
    If Not PropertyCalculation.HasElements(members) Then Exit Function

    Dim i As Long
    For i = LBound(members) To UBound(members)
        If Not members(i) Is Nothing Then
            If members(i).IsCellElement Then
                If AnyPushableSourcePatternMatchesTarget(members(i).AsCellElement.Name, pushedTargets) Then
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

' True when sName matches a pushable Cell* source's pattern of at least one calc rule whose TargetProp is
' in pushedTargets - the real competing-trigger predicate (see GroupHasCompetingTrigger).
Private Function AnyPushableSourcePatternMatchesTarget(ByVal sName As String, ByVal pushedTargets As Collection) As Boolean
    On Error GoTo ErrorHandler

    AnyPushableSourcePatternMatchesTarget = False
    Dim i As Long
    Dim rule As CalcRuleInfo
    For i = 0 To PropertyCalculation.RuleCount() - 1
        rule = PropertyCalculation.GetRule(i)
        If IsPushableCellSourceKind(rule.SourceKind) Then
            If PropertyCalculation_SourceEval.MatchesAnyPattern(sName, rule.SourceArg) Then
                If CollectionContainsString(pushedTargets, rule.TargetProp) Then
                    AnyPushableSourcePatternMatchesTarget = True
                    Exit Function
                End If
            End If
        End If
    Next i
    Exit Function

ErrorHandler:
    AnyPushableSourcePatternMatchesTarget = False
End Function

' Linear membership check on a Collection of strings (dedup helper - Collection has no native Contains).
' Case-insensitive (vbTextCompare), matching FindCalcRuleForProperty/DistinctTargets's TargetProp comparisons
' elsewhere in the engine - a calc rule's target-property name is never case-sensitive.
Private Function CollectionContainsString(ByVal col As Collection, ByVal s As String) As Boolean
    Dim v As Variant
    CollectionContainsString = False
    For Each v In col
        If StrComp(CStr(v), s, vbTextCompare) = 0 Then
            CollectionContainsString = True
            Exit Function
        End If
    Next v
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
    ' Target properties oTriggerEl actually pushed a value to (post AC3 guard) - the only props for which
    ' another trigger element in the group can be a REAL competitor. See GroupHasCompetingLvlTrigger.
    Dim pushedTargets As New Collection
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
                                If Not CollectionContainsString(pushedTargets, P) Then pushedTargets.Add P
                            End If
                        End If
                    End If
                End If
            Next ri
        End If
    Next i

    ' Discoverability: siblings match a rule but NONE carry its target -> attach never happened.
    If nCarried = 0 Then PropertyCalculation.ReportNoTarget

    ' Multi-trigger: another LEVEL-matching element in the group also feeds one of oTriggerEl's OWN pushed
    ' targets (last-processed wins). Restricted to pushedTargets - two elements feeding two DIFFERENT
    ' properties are not competitors. Own DISTINCT status (ReportMultipleLvlTriggers, not
    ' ReportMultipleTriggers): this collision may be pure Lvl-vs-Lvl, involving no cell at all - the message
    ' must not claim "cells". See the KNOWN LIMITATION note above the section header - cross-family
    ' (Cell-vs-Lvl) competition is still not detected here (a distinct, accepted gap, not this wording fix).
    If GroupHasCompetingLvlTrigger(oTriggerEl, pushedTargets) Then PropertyCalculation.ReportMultipleLvlTriggers
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_TriggerPush.PushLvlDerivedValuesToMembers"
End Sub

' True when oTriggerEl's graphic group holds at least one OTHER element whose Level ALSO matches a
' pushable Lvl* source feeding one of pushedTargets - the properties oTriggerEl itself just pushed a value
' to. Restricted to pushedTargets (not "any pushable rule") so that two elements feeding two DIFFERENT
' properties are not reported as competing - only a real collision on the SAME target property is. Mirrors
' GroupHasCompetingTrigger. NO element-type restriction (mirrors IsTriggerLevel).
Private Function GroupHasCompetingLvlTrigger(ByVal oTriggerEl As element, ByVal pushedTargets As Collection) As Boolean
    On Error GoTo ErrorHandler

    GroupHasCompetingLvlTrigger = False
    If pushedTargets.Count = 0 Then Exit Function

    Dim members() As element
    members = Link.GetLink(oTriggerEl)             ' OTHER members only
    If Not PropertyCalculation.HasElements(members) Then Exit Function

    Dim i As Long
    For i = LBound(members) To UBound(members)
        If Not members(i) Is Nothing Then
            If members(i).IsGraphical Then
                If Not members(i).Level Is Nothing Then
                    If AnyPushableLvlSourcePatternMatchesTarget(members(i).Level.Name, pushedTargets) Then
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

' True when sLevelName matches a pushable Lvl* source's pattern of at least one calc rule whose TargetProp
' is in pushedTargets - the real competing-trigger predicate (see GroupHasCompetingLvlTrigger).
Private Function AnyPushableLvlSourcePatternMatchesTarget(ByVal sLevelName As String, ByVal pushedTargets As Collection) As Boolean
    On Error GoTo ErrorHandler

    AnyPushableLvlSourcePatternMatchesTarget = False
    Dim i As Long
    Dim rule As CalcRuleInfo
    For i = 0 To PropertyCalculation.RuleCount() - 1
        rule = PropertyCalculation.GetRule(i)
        If IsPushableLvlSourceKind(rule.SourceKind) Then
            If PropertyCalculation_SourceEval.MatchesAnyPattern(sLevelName, rule.SourceArg) Then
                If CollectionContainsString(pushedTargets, rule.TargetProp) Then
                    AnyPushableLvlSourcePatternMatchesTarget = True
                    Exit Function
                End If
            End If
        End If
    Next i
    Exit Function

ErrorHandler:
    AnyPushableLvlSourcePatternMatchesTarget = False
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
