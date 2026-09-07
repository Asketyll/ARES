' Module: SheetLevels
' Description: Turns every level on, in every open view, of every Sheet ("Papier") model of the active
'              design file whose name matches ARES_Sheet_Levels_Model_Name.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, ARESConstants, Config, ErrorHandlerClass, LangManager, RuleGrammar
'
' WHICH SWITCH, AND WHY IT COSTS A MODEL ACTIVATION
' MicroStation paints a level only when Global Display is on AND the level is not frozen AND view display
' is on for that view (IsDisplayed / IsDisplayedInView Remarks, mvba-docs). Measured live on the target
' 24-folio file (2026-09-07): Global Display was ALREADY True on every level of every folio, which is why
' the first build reported "0 level(s) switched on" while the sheets showed nothing. The switch that was
' actually off is the PER-VIEW one - `IsDisplayedInView`, the value the Level Display dialog box drives.
' `IsDisplayed` is still written here, as the documented prerequisite, but on this file it is a no-op.
'
' The per-view masks are reachable only through `ActiveDesignFile.Views`, which is the ACTIVE view group's
' collection (Views_Object: "use the property, DesignFile.Views"), and `ViewGroup` exposes nothing but
' Name/Description/IsActive - there is no way to walk the views of a model that is not active. So each
' matching sheet is made active in turn (`ModelReference.Activate`), written, and the model that was
' active on entry is restored at the end - including when the walk faults. That model switching is the
' price of the documented mechanism; an earlier build wrote `Level.IsActive` instead, which does make a
' level visible (an active level cannot be masked) but only as a SIDE EFFECT, and it silently changed the
' active level of all 24 folios. Do not reinstate it.
'
' `IsFrozen` is deliberately still NOT written: a frozen level is an editorial decision the sheet's author
' made, and nothing in the request asked to undo it. A frozen level therefore stays invisible; run
' DiagSheetLevels to see how many there are.
'
' MESSAGE CHANNEL - everything this module refuses is an expected user or environment situation (no file
' open, an unconfigured pattern, a read-only model, no matching sheet), so it is reported on the status
' bar, translated. Only the absent design file ALSO writes one informational log line, mirroring
' CableReport's "No active model reference". Real faults - a level table that will not open or commit -
' go to the log in English and never abort the walk.
Option Explicit

' MicroStation's fixed view count. Views are addressed by index because the per-view mask is per view
' NUMBER; a closed view is skipped rather than opened.
Private Const MAX_VIEWS As Long = 8

' Sole public entry, driven by the key-in Command.ActivateSheetLevels. Walks the active design file's
' top-level models, keeps the Sheet ones whose NAME matches the configured pattern, and turns every level
' of each writable one on in every open view of that model.
Public Sub ActivateLevels()
    On Error GoTo ErrorHandler

    Dim oPrevModel As ModelReference

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

    ' Captured BEFORE the first Activate, restored on every exit path including the fault one.
    If Application.HasActiveModelReference Then Set oPrevModel = ActiveModelReference

    Dim oModel    As ModelReference
    Dim matched() As ModelReference
    Dim i         As Long
    Dim nModels   As Long   ' matching sheet models actually processed
    Dim nViews    As Long   ' open views written across all of them
    Dim nSwitched As Long   ' individual level switches turned on AND committed
    Dim nReadOnly As Long   ' matching sheet models skipped because they are read-only

    ' Pass 1 - COLLECT ONLY, no Activate. The writing pass makes models active, and mutating the active
    ' model from inside a For Each over the very collection being enumerated is not something MVBA
    ' documents as safe; collecting first removes the question entirely.
    ' Cheapest test first: the model type, then the name, and only then the model itself. Nested Ifs
    ' rather than one And, because VBA never short-circuits.
    ' The read-only test is per MODEL, the scope the write actually targets: a model can be read-only
    ' because it is locked as well as because the file is (IsReadOnly Remarks, mvba-docs), and
    ' OpenDesignFileForProgram's Remarks name ModelReference.IsReadOnly as the way to ask.
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
    For i = 0 To nModels - 1
        nSwitched = nSwitched + DisplayAllLevels(matched(i), nViews)
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

    ' Only when something actually moved. The restored model's own views keep painting the level set they
    ' had before this ran, so without the redraw a successful run can still look like it did nothing.
    If nSwitched > 0 Then RedrawAllViews

    If nReadOnly > 0 Then
        ShowStatus GetTranslation("SheetLevelsCompletePartial", nModels, nViews, nSwitched, nReadOnly)
    Else
        ShowStatus GetTranslation("SheetLevelsComplete", nModels, nViews, nSwitched)
    End If
    Exit Sub

ErrorHandler:
    ' The user's own model comes back before anything else - leaving them parked on folio 17 because a
    ' level table faulted would be a worse outcome than the fault itself.
    RestoreModel oPrevModel
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

' Turns every level of ONE model on, in each of that model's open views; returns how many switches moved
' AND were committed, and adds the views it wrote to nViews.
' The model is ACTIVATED first: ActiveDesignFile.Views is the active view group's collection, so a model's
' own views are unreachable until it is the active one (see the module header).
' Compare-before-write on every switch, so a sheet already fully on costs no write and no Rewrite.
' The count is reported only once the commit succeeds: an uncommitted level change is discarded when the
' design file closes (Rewrite Method Remarks, mvba-docs), so counting it would tell the user something
' untrue. Levels.Rewrite is called on the CACHED collection, not on a fresh oModel.Levels accessor - the
' doc's own example caches it, and a second accessor may hand back another wrapper whose Rewrite commits
' nothing. It is the per-model call rather than DesignFile.RewriteLevels because the latter acts on
' DesignFile.Levels, which is the DEFAULT model reference's collection (Levels Property Remarks).
' Its own error handler is what keeps one faulting model from aborting the whole run.
Private Function DisplayAllLevels(ByVal oModel As ModelReference, ByRef nViews As Long) As Long
    On Error GoTo ErrorHandler

    Dim oLevels  As Levels
    Dim oLevel   As Level
    Dim oView    As View
    Dim i        As Long
    Dim nChanged As Long

    ' Confirm the switch actually happened before reading ActiveDesignFile.Views: if Activate quietly did
    ' nothing, those views still belong to the PREVIOUS model, and writing their masks onto this model's
    ' levels would corrupt the wrong pairing. Informational (Number 0) so it does not turn a partial run
    ' into a generic "command failed"; the model simply contributes nothing.
    oModel.Activate
    If Not oModel.IsActive Then
        ErrorHandler.HandleError "Model did not become active, skipped: " & oModel.Name, 0, "", _
                                 "SheetLevels.DisplayAllLevels"
        Exit Function
    End If

    Set oLevels = oModel.Levels
    If oLevels Is Nothing Then Exit Function

    ' Global Display first, once per level: it is the documented prerequisite, it is not per-view, and
    ' re-testing it inside the view loop would read it MAX_VIEWS times for nothing.
    ' No Levels.Count guard: Count is undocumented for this collection, and an empty one never loops.
    For Each oLevel In oLevels
        If Not oLevel.IsDisplayed Then
            oLevel.IsDisplayed = True
            nChanged = nChanged + 1
        End If
    Next

    ' Then the per-view mask, view by view. Closed views are skipped rather than opened - the request is
    ' to show the levels of the sheets, not to change which views a sheet opens with.
    For i = 1 To MAX_VIEWS
        Set oView = GetOpenView(i)
        If Not oView Is Nothing Then
            nViews = nViews + 1
            For Each oLevel In oLevels
                If Not oLevel.IsDisplayedInView(oView) Then
                    oLevel.IsDisplayedInView(oView) = True
                    nChanged = nChanged + 1
                End If
            Next
        End If
    Next i

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

' The active view group's view at nIndex when it is OPEN, Nothing otherwise. Its own frame with Resume
' Next because an index that does not resolve is not a fault - it is simply a view that is not there,
' and it must not abort the model's pass.
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

' Puts back the model that was active before the walk. Own frame with Resume Next because it is also
' called from ActivateLevels' ErrorHandler, where that procedure's handler is already active and an
' inline On Error Resume Next would not reliably be in force. Nothing = nothing to restore.
Private Sub RestoreModel(ByVal oModel As ModelReference)
    On Error Resume Next

    If oModel Is Nothing Then Exit Sub
    oModel.Activate
    Err.Clear
End Sub

' Read-only measurement, writes nothing and activates nothing. For every matching sheet model it reports
' how many of its levels are OFF on each axis it can see without activating: global display and freeze.
' The PER-VIEW mask - the one ActivateLevels actually writes - is only reachable for the model that is
' already active, so it is reported for that one alone. It exists because ActivateLevels' own
' "0 switch(es)" is ambiguous: it means the same thing whether nothing was off in the first place or
' every commit silently failed, and SafeRewrite deliberately does not log.
' Output goes to the .log, in English, one line per model, plus the Immediate window - same shape as the
' CableReport diagnostics. Key-in: DiagSheetLevels.
Public Sub DiagnoseLevels()
    On Error GoTo ErrorHandler

    If Not ARESConfig.IsInitialized Then Exit Sub
    If Not Application.HasActiveDesignFile Then
        DiagLine "DIAG SheetLevels: no active design file."
        Exit Sub
    End If

    Dim sPattern As String
    sPattern = ResolvePattern()
    DiagLine "DIAG SheetLevels: pattern=[" & sPattern & "]"

    Dim oModel As ModelReference
    For Each oModel In ActiveDesignFile.Models
        If oModel.Type = msdModelTypeSheet Then
            If RuleGrammar.LikeAnyInListCI(oModel.Name, sPattern) Then DiagModel oModel
        End If
    Next
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels.DiagnoseLevels"
End Sub

' One measured line for one model. Counts are taken with no write, no Rewrite and no Activate.
Private Sub DiagModel(ByVal oModel As ModelReference)
    On Error GoTo ErrorHandler

    Dim oLevels    As Levels
    Dim oLevel     As Level
    Dim nTotal     As Long
    Dim nGlobalOff As Long
    Dim nFrozen    As Long

    Set oLevels = oModel.Levels
    If oLevels Is Nothing Then
        DiagLine "  " & oModel.Name & " | LEVELS COLLECTION IS NOTHING"
        Exit Sub
    End If

    For Each oLevel In oLevels
        nTotal = nTotal + 1
        If Not oLevel.IsDisplayed Then nGlobalOff = nGlobalOff + 1
        If oLevel.IsFrozen Then nFrozen = nFrozen + 1
    Next

    DiagLine "  " & oModel.Name & _
             " | modelActive=" & CStr(oModel.IsActive) & _
             " | readOnly=" & CStr(oModel.IsReadOnly) & _
             " | levels=" & CStr(nTotal) & _
             " | globalDisplayOFF=" & CStr(nGlobalOff) & _
             " | frozen=" & CStr(nFrozen)

    ' Per-view display is only measurable on the ACTIVE model: ActiveDesignFile.Views is the active view
    ' group's, and MVBA offers no way to reach another model's without activating it - which a read-only
    ' diagnostic must not do.
    If oModel.IsActive Then DiagActiveViews oLevels
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels.DiagModel"
End Sub

' Per-view "off" counts for the active model's open views, one line per view.
Private Sub DiagActiveViews(ByVal oLevels As Levels)
    On Error GoTo ErrorHandler

    Dim oView    As View
    Dim oLevel   As Level
    Dim i        As Long
    Dim nViewOff As Long

    For i = 1 To MAX_VIEWS
        Set oView = GetOpenView(i)
        If Not oView Is Nothing Then
            nViewOff = 0
            For Each oLevel In oLevels
                If Not oLevel.IsDisplayedInView(oView) Then nViewOff = nViewOff + 1
            Next
            DiagLine "      view " & CStr(i) & " (open) | viewDisplayOFF=" & CStr(nViewOff)
        End If
    Next i
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels.DiagActiveViews"
End Sub

' Diagnostic output: the .log (so it can be copied out of a session) and the Immediate window.
Private Sub DiagLine(ByVal sText As String)
    On Error Resume Next
    Debug.Print sText
    ErrorHandler.HandleError sText, 0, "", "SheetLevels.Diag"
End Sub

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
