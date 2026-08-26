' Module: PropertyRendering_Reporting
' Description: One-shot status reporters (status bar only, translated) plus the two session-scoped
'              self-disable log lines. Every refusal is EXPECTED user feedback - status-bar only, never
'              logged - except the schema/library self-disable conditions, the single mandated exception:
'              a feature that turned itself off must be greppable. Full mechanism: see
'              _bmad/docs/property-rendering-mechanics.md.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass (global ErrorHandler), PropertyRendering_TemplateModel

Option Explicit

' One-shot status guards - each refusal surfaces once per PROCESSED ELEMENT, reset in ResetOneShots.
' Same shape as PropertyCalculation's mbRejectedShown / mbNoTargetShown / mbMultiShown.
Private mbTokenUnknownShown As Boolean
Private mbValueUnsupportedShown As Boolean
Private mbValueIllegalShown As Boolean
Private mbMetadataInvalidShown As Boolean
Private mbMetadataUnreadableShown As Boolean
Private mbSchemaShown As Boolean
Private mbLibraryMissingShown As Boolean
Private mbAmbiguousShown As Boolean
Private mbBindingReleasedShown As Boolean

Private mbLockedShown As Boolean
Private mbDriftShown As Boolean
Private mbDuplicateShown As Boolean
Private mbAdjacentShown As Boolean
Private mbTextNodeRefusedShown As Boolean
Private mbNotBoundShown As Boolean
Private mbGovernedShown As Boolean
Private mbCycleShown As Boolean

' The two self-disable conditions also write an English log line. Those flags are SESSION-scoped, not
' per-element: ResetOneShots must not clear them, or a station whose DGNLib predates epic 15 would log
' one line per processed element for the whole session.
Private mbLibraryLogged As Boolean
Private mbSchemaLogged As Boolean

Public Sub ResetOneShots()
    On Error Resume Next
    mbTokenUnknownShown = False
    mbValueUnsupportedShown = False
    mbValueIllegalShown = False
    mbMetadataInvalidShown = False
    mbMetadataUnreadableShown = False
    mbSchemaShown = False
    mbLibraryMissingShown = False
    mbAmbiguousShown = False
    mbBindingReleasedShown = False
    mbLockedShown = False
    mbDriftShown = False
    mbDuplicateShown = False
    mbAdjacentShown = False
    mbTextNodeRefusedShown = False
    mbNotBoundShown = False
    mbGovernedShown = False
    mbCycleShown = False
End Sub

Public Sub ReportTokenUnknown()
    On Error Resume Next
    If Not mbTokenUnknownShown Then
        LangManager.ShowStatusT "RenderTokenUnknown"
        mbTokenUnknownShown = True
    End If
End Sub

Public Sub ReportValueUnsupported()
    On Error Resume Next
    If Not mbValueUnsupportedShown Then
        LangManager.ShowStatusT "RenderValueUnsupported"
        mbValueUnsupportedShown = True
    End If
End Sub

Public Sub ReportValueIllegal()
    On Error Resume Next
    If Not mbValueIllegalShown Then
        LangManager.ShowStatusT "RenderValueIllegalChars"
        mbValueIllegalShown = True
    End If
End Sub

Public Sub ReportMetadataInvalid()
    On Error Resume Next
    If Not mbMetadataInvalidShown Then
        LangManager.ShowStatusT "RenderMetadataInvalid"
        mbMetadataInvalidShown = True
    End If
End Sub

Public Sub ReportMetadataUnreadable()
    On Error Resume Next
    If Not mbMetadataUnreadableShown Then
        LangManager.ShowStatusT "RenderMetadataUnreadable"
        mbMetadataUnreadableShown = True
    End If
End Sub

' Self-disable condition: status AND ONE English log line - the single mandated exception to the
' "expected refusals are never logged" rule, because a feature that turned itself off must be greppable.
' The log is gated by the same one-shot flag so a whole batch cannot flood the .log.
Public Sub ReportSchemaUnsupported()
    On Error Resume Next
    If Not mbSchemaLogged Then
        ErrorHandler.HandleError "Property rendering: unsupported ARES_Render schema version, metadata left untouched", 0, "", "PropertyRendering.ReadRenderMetadata"
        mbSchemaLogged = True
    End If
    If Not mbSchemaShown Then
        LangManager.ShowStatusT "RenderSchemaUnsupported"
        mbSchemaShown = True
    End If
End Sub

' Self-disable condition: status AND ONE English log line (see ReportSchemaUnsupported).
Public Sub ReportLibraryMissing()
    On Error Resume Next
    If Not mbLibraryLogged Then
        ErrorHandler.HandleError "Property rendering: internal ARES_SYS item type library not found, rendering disabled", 0, "", "PropertyRendering.IsSysLibraryPresent"
        mbLibraryLogged = True
    End If
    If Not mbLibraryMissingShown Then
        LangManager.ShowStatusT "RenderLibraryMissing"
        mbLibraryMissingShown = True
    End If
End Sub

Public Sub ReportAmbiguous()
    On Error Resume Next
    If Not mbAmbiguousShown Then
        LangManager.ShowStatusT "RenderAmbiguousEdit"
        mbAmbiguousShown = True
    End If
End Sub

Public Sub ReportBindingReleased()
    On Error Resume Next
    If Not mbBindingReleasedShown Then
        LangManager.ShowStatusT "RenderBindingReleased"
        mbBindingReleasedShown = True
    End If
End Sub

Public Sub ReportLocked()
    On Error Resume Next
    If Not mbLockedShown Then
        LangManager.ShowStatusT "RenderElementLocked"
        mbLockedShown = True
    End If
End Sub

Public Sub ReportDrift()
    On Error Resume Next
    If Not mbDriftShown Then
        LangManager.ShowStatusT "RenderSubIdDrift"
        mbDriftShown = True
    End If
End Sub

Public Sub ReportDuplicateToken()
    On Error Resume Next
    If Not mbDuplicateShown Then
        LangManager.ShowStatusT "RenderDuplicateToken"
        mbDuplicateShown = True
    End If
End Sub

Public Sub ReportAdjacentTokens()
    On Error Resume Next
    If Not mbAdjacentShown Then
        LangManager.ShowStatusT "RenderAdjacentTokens"
        mbAdjacentShown = True
    End If
End Sub

Public Sub ReportTextNodeRefused()
    On Error Resume Next
    If Not mbTextNodeRefusedShown Then
        LangManager.ShowStatusT "RenderTextNodeInCellRefused"
        mbTextNodeRefusedShown = True
    End If
End Sub

Public Sub ReportNotBound()
    On Error Resume Next
    If Not mbNotBoundShown Then
        LangManager.ShowStatusT "RenderNotBound"
        mbNotBoundShown = True
    End If
End Sub

' Report the exact structural reason a first author was refused (duplicate vs adjacent tokens).
Public Sub ReportStructuralRefusal(ByVal sTemplate As String)
    On Error Resume Next

    Dim lits() As String
    Dim toks() As String
    Dim nTok As Long
    Dim i As Long, j As Long

    If Not PropertyRendering_TemplateModel.ParseTemplate(sTemplate, lits, toks, nTok, True) Then Exit Sub

    For i = 0 To nTok - 1
        For j = i + 1 To nTok - 1
            If StrComp(toks(i), toks(j), vbTextCompare) = 0 Then
                ReportDuplicateToken
                Exit Sub
            End If
        Next j
    Next i

    ReportAdjacentTokens
End Sub

' Edit-time discoverability: tell the user a value they just edited is governed by a calc rule (and so the
' edit will not stick unless the token was released too). Called by Module D (StateMachine)'s
' WarnGovernedTokensLost, at the moment a user edit costs the binding a governed token - NOT at bind time
' (moved off Module E's original WarnGovernedValue, 2026-08-26, per Asketyll's field feedback: the message
' text reads as a response to an edit, so it must not fire before any edit happened). mbGovernedShown
' stays here per the "all one-shot flags in one place" judgment call (§1 option (a) of the split plan,
' decided by Asketyll via the lead).
Public Sub ReportGovernedValue(ByVal P As String)
    On Error Resume Next
    If Not mbGovernedShown Then
        ShowStatus LangManager.GetTranslation("RenderValueGoverned", P)
        mbGovernedShown = True
    End If
End Sub

' "This value is computed from the text of this very cell; the rendered text is excluded from that
' computation." Used both at bind time, as the static-cycle warning, and by the top-up scan when it
' refuses the cell's LAST unrendered sub-text - in both cases the message is exactly the point being made.
Public Sub ReportCycleWarning()
    On Error Resume Next
    If Not mbCycleShown Then
        LangManager.ShowStatusT "RenderCycleWarning"
        mbCycleShown = True
    End If
End Sub
