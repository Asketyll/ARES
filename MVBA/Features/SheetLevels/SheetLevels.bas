' Module: SheetLevels
' Description: Turns Global Display ON for every level of every Sheet ("Papier") model of the active
'              design file whose name matches ARES_Sheet_Levels_Model_Name.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, ARESConstants, Config, ErrorHandlerClass, LangManager, RuleGrammar
'
' SCOPE - deliberately one switch out of the three that gate a level's visibility. MicroStation paints
' a level only when Global Display is on AND the level is not frozen AND view display is on for that
' view (Level.IsDisplayed Remarks, mvba-docs). This module writes Global Display and nothing else:
'   - IsFrozen is left alone: a frozen level is an editorial decision the sheet's author made.
'   - IsDisplayedInView is left alone because it is not REACHABLE from here. The per-view level masks
'     belong to each model's own view group, and MVBA exposes no way to walk the views of a model that
'     is not the active one (ViewGroup carries Name/Description/IsActive and nothing else). Getting to
'     them means ModelReference.Activate on every matching sheet in turn, then restoring the user's own
'     model - a model switch, its view updates and a restore this key-in deliberately does not pay for.
' A level left frozen, or switched off in the sheet's own view, therefore stays invisible. By design.
'
' MESSAGE CHANNEL - everything this module refuses is an expected user or environment situation (no file
' open, an unconfigured pattern, a read-only model, no matching sheet), so it is reported on the status
' bar, translated. Only the absent design file ALSO writes one informational log line, mirroring
' CableReport's "No active model reference". Real faults - a level table that will not open or commit -
' go to the log in English and never abort the walk.
Option Explicit

' Sole public entry, driven by the key-in Command.ActivateSheetLevels. Walks the active design file's
' top-level models, keeps the Sheet ones whose NAME matches the configured pattern, and turns Global
' Display on for every level of each writable one.
Public Sub ActivateLevels()
    On Error GoTo ErrorHandler

    If Not ARESConfig.IsInitialized Then
        ErrorHandler.HandleError "ARESConfig not initialized", 0, "", "SheetLevels.ActivateLevels"
        Exit Sub
    End If

    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    ' ActiveDesignFile RAISES when nothing is open - the guard is HasActiveDesignFile, never an
    ' "Is Nothing" test, which would have to read the property to compare it.
    If Not Application.HasActiveDesignFile Then
        ErrorHandler.HandleError "No active design file", 0, "", "SheetLevels.ActivateLevels"
        ShowStatusT "SheetLevelsNoDesignFile"
        Exit Sub
    End If

    Dim sPattern As String
    sPattern = ResolvePattern()

    ' Fail-closed on an empty pattern. "Every sheet model in the file" is not a safe reading of an
    ' empty setting - it is an UNCONFIGURED one, and this key-in writes to models the user is not
    ' looking at. ResolvePattern normalises delimiters and blanks away, so "|", "||" and " | " all
    ' land here too rather than walking every model and matching none.
    If Len(sPattern) = 0 Then
        ShowStatusT "SheetLevelsPatternEmpty"
        Exit Sub
    End If

    Dim oModel    As ModelReference
    Dim nModels   As Long   ' matching sheet models actually processed
    Dim nLevels   As Long   ' levels switched on AND committed
    Dim nReadOnly As Long   ' matching sheet models skipped because they are read-only

    ' Cheapest test first: the model type, then the name, and only then the model itself - reaching a
    ' model's level collection makes MicroStation materialise its level table, which is the real cost
    ' on a file holding many sheets. Nested Ifs rather than one And, because VBA never short-circuits.
    ' The read-only test is per MODEL, the scope the write actually targets: a model can be read-only
    ' because it is locked as well as because the file is (IsReadOnly Remarks, mvba-docs), and
    ' OpenDesignFileForProgram's Remarks name ModelReference.IsReadOnly as the way to ask.
    For Each oModel In ActiveDesignFile.Models
        If oModel.Type = msdModelTypeSheet Then
            If RuleGrammar.LikeAnyInListCI(oModel.Name, sPattern) Then
                If oModel.IsReadOnly Then
                    nReadOnly = nReadOnly + 1
                Else
                    nLevels = nLevels + DisplayAllLevels(oModel)
                    nModels = nModels + 1
                End If
            End If
        End If
    Next

    If nModels = 0 Then
        If nReadOnly > 0 Then
            ShowStatus GetTranslation("SheetLevelsReadOnly", nReadOnly)
        Else
            ShowStatus GetTranslation("SheetLevelsNoModel", sPattern)
        End If
        Exit Sub
    End If

    ' Only when something actually moved. If one of the matched sheets IS the active model, its open
    ' views keep painting the old level set until a redraw - which is what makes a successful run look
    ' like it did nothing. It writes no data and activates no model, so it does not widen this key-in's
    ' scope (same closing move as the mvba-docs level-display example).
    If nLevels > 0 Then RedrawAllViews

    If nReadOnly > 0 Then
        ShowStatus GetTranslation("SheetLevelsCompletePartial", nModels, nLevels, nReadOnly)
    Else
        ShowStatus GetTranslation("SheetLevelsComplete", nModels, nLevels)
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels.ActivateLevels"
End Sub

' The configured model-name pattern, normalised: alternatives trimmed, blank ones dropped, rejoined on
' ARES_VAR_DELIMITER. "" means "not configured" - including a value that is nothing but delimiters.
' Trimming here rather than inside RuleGrammar.LikeAnyInListCI is deliberate: that helper is shared with
' PropertyCalculation, whose Cell*[pattern] arguments must keep their current untrimmed semantics. The
' result matches what a tag/calc CONDITION does with the same "name|name" list (ParseCondition trims).
' The value is read LIVE, not from ARESConfig's boot-time snapshot: this var has no options form, so the
' MicroStation Configuration dialog is the only way to change it, and a snapshot would ignore that edit
' until the next restart. Falls back to the snapshot when the variable is not defined at all.
Private Function ResolvePattern() As String
    On Error GoTo ErrorHandler

    Dim sRaw As String
    sRaw = Config.GetVar(ARESConfig.ARES_SHEET_LEVELS_MODEL_NAME.Key)
    If sRaw = ARESConstants.ARES_NAVD Then sRaw = ARESConfig.ARES_SHEET_LEVELS_MODEL_NAME.Value

    ' SplitTrim drops blank parts and returns a single "" when every part is blank.
    ResolvePattern = Join(RuleGrammar.SplitTrim(sRaw, ARESConstants.ARES_VAR_DELIMITER), _
                          ARESConstants.ARES_VAR_DELIMITER)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels.ResolvePattern"
    ResolvePattern = ""
End Function

' Turns Global Display on for every level of ONE model; returns how many levels moved AND were committed.
' Compare-before-write, so a sheet that is already fully displayed costs no write and no Rewrite.
' The count is reported only once the commit succeeds: an uncommitted level change is discarded when the
' design file closes (Rewrite Method Remarks, mvba-docs), so counting it would tell the user something
' untrue. Levels.Rewrite is called on the CACHED collection, not on a fresh oModel.Levels accessor - the
' doc's own example caches it, and a second accessor may hand back another wrapper whose Rewrite commits
' nothing. It is the per-model call rather than DesignFile.RewriteLevels because the latter acts on
' DesignFile.Levels, which is the DEFAULT model reference's collection (Levels Property Remarks).
' Its own error handler is what keeps one faulting model from aborting the whole run.
Private Function DisplayAllLevels(ByVal oModel As ModelReference) As Long
    On Error GoTo ErrorHandler

    Dim oLevels  As Levels
    Dim oLevel   As Level
    Dim nChanged As Long

    Set oLevels = oModel.Levels
    If oLevels Is Nothing Then Exit Function

    ' No Levels.Count guard: Count is undocumented for this collection, and an empty one simply never
    ' enters the loop.
    For Each oLevel In oLevels
        If Not oLevel.IsDisplayed Then
            oLevel.IsDisplayed = True
            nChanged = nChanged + 1
        End If
    Next

    If nChanged = 0 Then Exit Function
    If SafeRewrite(oLevels) Then DisplayAllLevels = nChanged
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels.DisplayAllLevels"
    ' Best effort: commit whatever landed before the fault rather than leaving the model changed in
    ' memory and uncommitted on disk. Still counted only if that commit reports success.
    If nChanged > 0 Then
        If SafeRewrite(oLevels) Then DisplayAllLevels = nChanged
    End If
End Function

' Commits a level collection, reporting whether it worked. It lives in its OWN procedure on purpose:
' it is also called from DisplayAllLevels' ErrorHandler block, where that procedure's handler is already
' active and an inline On Error Resume Next would not reliably be in force. A fresh frame never is.
Private Function SafeRewrite(ByVal oLevels As Levels) As Boolean
    On Error Resume Next

    Err.Clear
    oLevels.Rewrite
    SafeRewrite = (Err.Number = 0)
    Err.Clear
End Function
