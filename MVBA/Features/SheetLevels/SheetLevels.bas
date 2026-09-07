' Module: SheetLevels
' Description: Turns every level on, in every open view, of every Sheet ("Papier") model of the active
'              design file whose name matches ARES_Sheet_Levels_Model_Name - and of everything those
'              sheets reference. Which switch, why each sheet is activated, and what was measured:
'              _bmad/docs/sheet-levels-mechanics.md
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, ARESConstants, Config, ErrorHandlerClass, LangManager, RuleGrammar
Option Explicit

' MicroStation's fixed view count; the per-view mask is per view NUMBER.
Private Const MAX_VIEWS As Long = 8

' Depth cap on the reference tree - a reference can itself reference, with no cheap loop detection here.
Private Const MAX_ATTACH_DEPTH As Long = 8

' Sole public entry, driven by the key-in Command.ActivateSheetLevels.
Public Sub ActivateLevels()
    On Error GoTo ErrorHandler

    Dim oPrevModel As ModelReference

    If Not ARESConfig.IsInitialized Then
        ErrorHandler.HandleError "ARESConfig not initialized", 0, "", "SheetLevels.ActivateLevels"
        Exit Sub
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    ' HasActiveDesignFile, not an "Is Nothing" test: reading ActiveDesignFile with nothing open RAISES.
    If Not Application.HasActiveDesignFile Then
        ErrorHandler.HandleError "No active design file", 0, "", "SheetLevels.ActivateLevels"
        ShowStatusT "SheetLevelsNoDesignFile"
        Exit Sub
    End If

    Dim sPattern As String
    sPattern = ResolvePattern()
    If Len(sPattern) = 0 Then
        ShowStatusT "SheetLevelsPatternEmpty"
        Exit Sub
    End If

    ' Captured BEFORE the first Activate, restored on every exit path including the fault one.
    If Application.HasActiveModelReference Then Set oPrevModel = ActiveModelReference

    Dim oModel    As ModelReference
    Dim matched() As ModelReference
    Dim i         As Long
    Dim nModels   As Long   ' matching sheet models actually processed
    Dim nViews    As Long   ' open views written across all of them
    Dim nSwitched As Long   ' individual level switches turned on AND committed
    Dim nReadOnly As Long   ' matching sheet models skipped because they are read-only

    ' Pass 1 - collect only. Pass 2 activates models, which must not happen while this enumeration runs.
    ' Cheapest test first; nested Ifs rather than one And, because VBA never short-circuits.
    ReDim matched(0 To 15)
    For Each oModel In ActiveDesignFile.Models
        If oModel.Type = msdModelTypeSheet Then
            If RuleGrammar.LikeAnyInListCI(oModel.Name, sPattern) Then
                If oModel.IsReadOnly Then
                    nReadOnly = nReadOnly + 1
                Else
                    If nModels > UBound(matched) Then ReDim Preserve matched(0 To UBound(matched) * 2 + 1)
                    Set matched(nModels) = oModel
                    nModels = nModels + 1
                End If
            End If
        End If
    Next

    ' Pass 2 - activate and write, one model at a time.
    Dim bWithAttachments As Boolean
    bWithAttachments = ProcessAttachments()

    For i = 0 To nModels - 1
        nSwitched = nSwitched + DisplayAllLevels(matched(i), nViews, bWithAttachments)
    Next i

    RestoreModel oPrevModel

    If nModels = 0 Then
        If nReadOnly > 0 Then
            ShowStatus GetTranslation("SheetLevelsReadOnly", nReadOnly)
        Else
            ShowStatus GetTranslation("SheetLevelsNoModel", sPattern)
        End If
        Exit Sub
    End If

    ' Without this the restored model keeps painting the level set it had on entry.
    If nSwitched > 0 Then RedrawAllViews

    If nReadOnly > 0 Then
        ShowStatus GetTranslation("SheetLevelsCompletePartial", nModels, nViews, nSwitched, nReadOnly)
    Else
        ShowStatus GetTranslation("SheetLevelsComplete", nModels, nViews, nSwitched)
    End If
    Exit Sub

ErrorHandler:
    RestoreModel oPrevModel   ' before anything else - never leave the user parked on folio 17
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels.ActivateLevels"
End Sub

' The configured model-name pattern, normalised: alternatives trimmed, blanks dropped, rejoined.
' "" = not configured, which includes a value that is nothing but delimiters. Read live, not from
' ARESConfig's snapshot, so an edit in the MicroStation Configuration dialog applies without a restart.
Private Function ResolvePattern() As String
    On Error GoTo ErrorHandler

    Dim sRaw As String
    sRaw = Config.GetVar(ARESConfig.ARES_SHEET_LEVELS_MODEL_NAME.Key)
    If sRaw = ARESConstants.ARES_NAVD Then sRaw = ARESConfig.ARES_SHEET_LEVELS_MODEL_NAME.Value

    ' SplitTrim drops blank parts and returns a single "" when every part is blank. Trimming here rather
    ' than in LikeAnyInListCI: that helper is shared with PropertyCalculation, whose Cell*[pattern]
    ' arguments must keep their untrimmed semantics.
    ResolvePattern = Join(RuleGrammar.SplitTrim(sRaw, ARESConstants.ARES_VAR_DELIMITER), _
                          ARESConstants.ARES_VAR_DELIMITER)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels.ResolvePattern"
    ResolvePattern = ""
End Function

' True when a sheet's references must be processed too (ARES_Sheet_Levels_Attachments). Anything that is
' not literally "False" - including an unreadable value - is True: doing half the job silently is worse.
Private Function ProcessAttachments() As Boolean
    On Error GoTo ErrorHandler

    Dim sRaw As String
    sRaw = Config.GetVar(ARESConfig.ARES_SHEET_LEVELS_ATTACHMENTS.Key)
    If sRaw = ARESConstants.ARES_NAVD Then sRaw = ARESConfig.ARES_SHEET_LEVELS_ATTACHMENTS.Value

    ProcessAttachments = (UCase(Trim(sRaw)) <> "FALSE")
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels.ProcessAttachments"
    ProcessAttachments = True
End Function

' Processes ONE sheet model: activates it, collects its open views, writes its levels and its references'.
' Returns the switches that moved AND were committed; adds the views it wrote to nViews.
' Its own error handler is what keeps one faulting model from aborting the whole run.
Private Function DisplayAllLevels(ByVal oModel As ModelReference, ByRef nViews As Long, _
                                  ByVal bWithAttachments As Boolean) As Long
    On Error GoTo ErrorHandler

    Dim oViews() As View
    Dim nOpen    As Long
    Dim nChanged As Long

    ' If Activate quietly did nothing, ActiveDesignFile.Views still holds the PREVIOUS model's views and
    ' their masks would be written onto this model's levels. Informational, so a skip is not a failure.
    oModel.Activate
    If Not oModel.IsActive Then
        ErrorHandler.HandleError "Model did not become active, skipped: " & oModel.Name, 0, "", _
                                 "SheetLevels.DisplayAllLevels"
        Exit Function
    End If

    nOpen = CollectOpenViews(oViews)
    nViews = nViews + nOpen

    ' A reference carries its OWN Levels collection: the sheet's levels say nothing about what it shows.
    nChanged = TurnOnLevels(oModel.Levels, oViews, nOpen)
    If bWithAttachments Then nChanged = nChanged + TurnOnAttachments(oModel, oViews, nOpen, 1)

    DisplayAllLevels = nChanged
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels.DisplayAllLevels"
    DisplayAllLevels = nChanged
End Function

' The active view group's open views, gathered once per model. Fills oViews from index 0, returns how many.
Private Function CollectOpenViews(ByRef oViews() As View) As Long
    On Error GoTo ErrorHandler

    Dim oView As View
    Dim i     As Long
    Dim n     As Long

    ReDim oViews(0 To MAX_VIEWS - 1)
    For i = 1 To MAX_VIEWS
        Set oView = GetOpenView(i)
        If Not oView Is Nothing Then
            Set oViews(n) = oView
            n = n + 1
        End If
    Next i

    CollectOpenViews = n
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels.CollectOpenViews"
    CollectOpenViews = n
End Function

' Turns every level of ONE Levels collection on: global display once per level, then the per-view mask in
' each open view. Compare-before-write throughout. Counted only once the Rewrite lands - an uncommitted
' level change is discarded when the file closes.
' Takes the collection, not the model reference: an Attachment is a distinct COM interface from
' ModelReference, and this avoids needing a conversion between them.
Private Function TurnOnLevels(ByVal oLevels As Levels, ByRef oViews() As View, ByVal nOpen As Long) As Long
    On Error GoTo ErrorHandler

    Dim oLevel   As Level
    Dim i        As Long
    Dim nChanged As Long

    If oLevels Is Nothing Then Exit Function

    ' No Levels.Count guard: Count is undocumented for this collection, and an empty one never loops.
    For Each oLevel In oLevels
        If Not oLevel.IsDisplayed Then
            oLevel.IsDisplayed = True
            nChanged = nChanged + 1
        End If
        For i = 0 To nOpen - 1
            If Not oLevel.IsDisplayedInView(oViews(i)) Then
                oLevel.IsDisplayedInView(oViews(i)) = True
                nChanged = nChanged + 1
            End If
        Next i
    Next

    If nChanged = 0 Then Exit Function
    If SafeRewrite(oLevels) Then TurnOnLevels = nChanged
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels.TurnOnLevels"
    ' Commit whatever landed before the fault; still counted only if it reports success.
    If nChanged > 0 Then
        If SafeRewrite(oLevels) Then TurnOnLevels = nChanged
    End If
End Function

' Depth-first walk of a model reference's attachments, turning each one's levels on.
' oRef is As Object so the same routine takes a ModelReference (the sheet) and an Attachment (a nested
' reference) - separate COM interfaces VBA would not convert between.
' An attachment is NEVER activated (it cannot be), and its IsReadOnly is NOT a skip reason - that flag is
' about its elements, its level display is writable. A missing file/model IS skipped: most of an
' Attachment's members raise once the file is gone.
Private Function TurnOnAttachments(ByVal oRef As Object, ByRef oViews() As View, _
                                   ByVal nOpen As Long, ByVal nDepth As Long) As Long
    On Error GoTo ErrorHandler

    Dim oAtt     As Attachment
    Dim nChanged As Long

    If nDepth > MAX_ATTACH_DEPTH Then Exit Function

    For Each oAtt In oRef.Attachments
        If Not oAtt.IsMissingFile Then
            If Not oAtt.IsMissingModel Then
                nChanged = nChanged + TurnOnLevels(oAtt.Levels, oViews, nOpen)
                nChanged = nChanged + TurnOnAttachments(oAtt, oViews, nOpen, nDepth + 1)
            End If
        End If
    Next

    TurnOnAttachments = nChanged
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels.TurnOnAttachments"
    TurnOnAttachments = nChanged
End Function

' The active view group's view at nIndex when it is OPEN, Nothing otherwise. An index that does not
' resolve is a view that is not there, not a fault.
Private Function GetOpenView(ByVal nIndex As Long) As View
    On Error Resume Next

    Dim oView As View

    Err.Clear
    Set oView = ActiveDesignFile.Views(nIndex)
    If Err.Number = 0 Then
        If Not oView Is Nothing Then
            If oView.IsOpen Then Set GetOpenView = oView
        End If
    End If
    Err.Clear
End Function

' Puts back the model that was active before the walk. Own procedure: it is called from ActivateLevels'
' ErrorHandler, where an inline On Error Resume Next would not reliably be in force. A fresh frame is.
Private Sub RestoreModel(ByVal oModel As ModelReference)
    On Error Resume Next

    If oModel Is Nothing Then Exit Sub
    oModel.Activate
    Err.Clear
End Sub

' Commits a level collection, reporting whether it worked. Own procedure for the same reason as
' RestoreModel: it is also called from an already-active error handler.
Private Function SafeRewrite(ByVal oLevels As Levels) As Boolean
    On Error Resume Next

    Err.Clear
    oLevels.Rewrite
    SafeRewrite = (Err.Number = 0)
    Err.Clear
End Function
