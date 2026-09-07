' Module: SheetLevels
' Description: Turns every level on, in every open view, of every Sheet ("Papier") model of the active
'              design file whose name matches ARES_Sheet_Levels_Model_Name - and of everything those
'              sheets reference.
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
' REFERENCES COUNT TOO. A sheet almost always shows its content through an attached reference (the design
' model), and a reference carries its OWN Levels collection with its own per-view mask - measured on the
' real file (2026-09-07): the folio's own levels came out right while the referenced model stayed blank.
' So each sheet's attachments are walked depth-first and given the same treatment, bounded by
' MAX_ATTACH_DEPTH since a reference can itself reference. An attachment is never activated (it cannot be)
' and its `IsReadOnly` is not a reason to skip it: that flag is about its ELEMENTS, while its level display
' is writable - which is precisely what the mvba-docs example Changing_Level_Display_for_an_Attachment
' does. Its levels need their own `Levels.Rewrite`; `DesignFile.RewriteLevels` explicitly does not reach
' them.
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

' Depth cap on the reference tree. A reference can itself reference, and there is no cheap identity test
' here to detect a chain that loops back, so the walk is bounded rather than trusted.
Private Const MAX_ATTACH_DEPTH As Long = 8

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

' Processes ONE sheet model: makes it active, collects its open views, then turns the levels on for the
' sheet itself AND for everything it references. Returns how many switches moved AND were committed, and
' adds the views it wrote to nViews.
' The model is ACTIVATED first: ActiveDesignFile.Views is the active view group's collection, so a model's
' own views are unreachable until it is the active one (see the module header).
' Its own error handler is what keeps one faulting model from aborting the whole run.
Private Function DisplayAllLevels(ByVal oModel As ModelReference, ByRef nViews As Long) As Long
    On Error GoTo ErrorHandler

    Dim oViews() As View
    Dim nOpen    As Long
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

    nOpen = CollectOpenViews(oViews)
    nViews = nViews + nOpen

    ' The sheet's own levels, then the levels of everything attached to it. A reference carries its OWN
    ' Levels collection, with its own per-view mask and its own Rewrite - turning the sheet's levels on
    ' says nothing about what its references show, which is exactly the gap found on the real file
    ' (2026-09-07: folio levels correct, referenced design model still blank).
    nChanged = TurnOnLevels(oModel.Levels, oViews, nOpen)
    nChanged = nChanged + TurnOnAttachments(oModel, oViews, nOpen, 1)

    DisplayAllLevels = nChanged
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels.DisplayAllLevels"
    DisplayAllLevels = nChanged
End Function

' The open views of the ACTIVE view group, gathered once per model so the level loop does not re-resolve
' them per level. Returns how many were found; oViews is filled from index 0.
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

' Turns every level of ONE Levels collection on - the sheet's own, or one of its references'. Global
' Display once per level (the documented prerequisite, not per-view), then the per-view mask in each open
' view. Compare-before-write on every switch, so a collection already fully on costs no write and no
' Rewrite. Closed views are not in oViews: the request is to show the levels of the sheets, not to change
' which views a sheet opens with.
' The count is returned only once the commit succeeds: an uncommitted level change is discarded when the
' design file closes (Rewrite Method Remarks, mvba-docs), so counting it would tell the user something
' untrue. Rewrite is called on the collection passed in, never on a fresh accessor - the doc's own example
' caches it, and a second accessor may hand back another wrapper whose Rewrite commits nothing. It is also
' the ONLY thing that commits an attachment's levels: the same Remarks state that DesignFile.RewriteLevels
' "does not rewrite the level information for attachments".
' Takes the Levels collection rather than the model reference on purpose: an Attachment is a distinct COM
' interface from ModelReference (hence ModelReference.AsAttachment), so passing one where the other is
' declared is a conversion this avoids needing.
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
    ' Best effort: commit whatever landed before the fault rather than leaving the collection changed in
    ' memory and uncommitted on disk. Still counted only if that commit reports success.
    If nChanged > 0 Then
        If SafeRewrite(oLevels) Then TurnOnLevels = nChanged
    End If
End Function

' Depth-first walk of a model reference's attachments, turning each one's levels on. A reference can
' itself reference, so it recurses, bounded by MAX_ATTACH_DEPTH.
' oRef is typed As Object so the same routine takes both a ModelReference (the sheet) and an Attachment
' (a nested reference) - they are separate COM interfaces and VBA would not convert one to the other.
' An attachment is NEVER activated: "an Attachment is a read-only ModelReference that cannot become the
' active model reference ... Activate raises an error" (Attachment_Object Remarks). Read-only there refers
' to its ELEMENTS - its level display is writable, which is what the doc's own
' Changing_Level_Display_for_an_Attachment example does. So the top-level IsReadOnly skip must NOT be
' applied here, or every reference would be silently passed over.
' A missing reference file is skipped rather than faulted through: "If the file is missing, many of the
' methods and properties of Attachment raise errors" (IsMissingFile Remarks), and one drawing with a
' broken reference would otherwise log a line per folio.
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
             " | frozen=" & CStr(nFrozen) & _
             " | attachments=" & CStr(CountAttachments(oModel))

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

' How many references a model carries, for the diagnostic line. Its own frame with Resume Next: a model
' whose attachments cannot be read must still produce its measurement line.
Private Function CountAttachments(ByVal oRef As Object) As Long
    On Error Resume Next

    Dim n As Long

    Err.Clear
    n = oRef.Attachments.Count
    If Err.Number = 0 Then CountAttachments = n
    Err.Clear
End Function

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
