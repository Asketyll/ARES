VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PropertyTagging_GUI_Options 
   Caption         =   "PropertyTagging_GUI_Options"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5310
   OleObjectBlob   =   "PropertyTagging_GUI_Options.frx":0000
End
Attribute VB_Name = "PropertyTagging_GUI_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm: PropertyTagging_GUI_Options
' Description: Options panel for Property Tagging - the master switch (ARES_Auto_Properties), the
'              custom-property list (ARES_Custom_Property_List, hidden reveal), and the attach rules
'              (ARES_Property_Rules). Rules are edited one at a time through an editable ComboBox
'              (split on ";"): pick a rule -> edit -> commit replaces it; free-type + commit appends;
'              empty + commit removes it. The ComboBox is the SOLE rules editor (the raw reveal was
'              removed - bulk config travels via .cfg import/export). Every commit is validated by
'              PropertyTagging.ValidateAndNormalizeRule, so a malformed rule (chiefly the "|"-instead-of-";"
'              mistake) is refused instead of saved.
'
'              Below ComboBox_Rules a read-only COLOURED SYNTAX PREVIEW (13-3) re-renders the current rule
'              via runtime Labels in a fixed-pitch font (an MSForms ComboBox cannot colour individual
'              characters): @ pink, & dark blue, ! orange, */? green, contradiction segments RED (via
'              PropertyTagging.RuleHasNoEffect), the rest monochrome. It renders the CANONICAL form when the
'              rule is valid (else the raw text, metachar colours only) and never writes anything.
'
'              DESIGNER (manual, Asketyll) - controls required with EXACTLY these names:
'                Main_CheckBox (CheckBox, master), Edit_PropertyList_Command (CommandButton) +
'                TextBox_PropertyList (TextBox, hidden reveal - property list), ComboBox_Rules (ComboBox,
'                Style = 0 fmStyleDropDownCombo EDITABLE - the sole per-rule editor),
'                Frame_RulePreview (Frame, render surface directly BELOW ComboBox_Rules for the runtime
'                coloured preview - the coloured Labels are created at runtime, NONE in the designer; resize
'                the form taller to fit it), Reset_Command (CommandButton).
'              StartUpPosition = 0 Manual. Tab order: master -> property-list -> rules-combo -> reset
'              (Frame_RulePreview is a non-focusable container, not in the tab order).
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass, ARESConfigClass, PropertyTagging, FormUXHelper, FormPlacement, Command
Option Explicit

' Rule separator for the editable ComboBox split/join. Mirrors PropertyTagging's own RULE_SEPARATOR
' (kept private there); the form only edits ARES_Property_Rules, it does not parse the grammar.
Private Const RULE_SEP As String = ";"

Private mbLocked As Boolean

' The ComboBox list index the user picked before editing its text (-1 = new / free-typed). Captured in
' _Change on a clean pick (Text = List(ListIndex)) because MSForms resets ListIndex to -1 once the edited
' text diverges from the selected item, so "which rule am I editing?" cannot be read at commit time.
Private mRuleEditIndex As Long

' --- Coloured syntax preview (13-3): read-only runtime Labels rendered under ComboBox_Rules ---
' Names of the runtime preview Labels currently on the form; each re-render removes exactly these before
' creating the new set, so the control count stays bounded (no leak across keystrokes/picks).
Private moPreviewLabels As Collection
' Monotonic sequence for unique Label names (never reset, so a failed Remove can never cause a name clash).
Private mlPreviewSeq As Long

' Preview colours (OLE_COLOR = RGB Long; tunable live). Render priority: contradiction RED > metachar > mono.
Private Const PREVIEW_PINK   As Long = 220 + 60 * 256& + 160 * 65536     ' @   group modifier    RGB(220,60,160)
Private Const PREVIEW_BLUE   As Long = 0 + 0 * 256& + 160 * 65536        ' &   AND               RGB(0,0,160)
Private Const PREVIEW_ORANGE As Long = 255 + 140 * 256& + 0 * 65536      ' !   negation          RGB(255,140,0)
Private Const PREVIEW_GREEN  As Long = 0 + 140 * 256& + 0 * 65536        ' */? wildcards         RGB(0,140,0)
Private Const PREVIEW_RED    As Long = 200 + 0 * 256& + 0 * 65536        ' contradiction segment RGB(200,0,0)
Private Const PREVIEW_MONO   As Long = 0                                 ' everything else       RGB(0,0,0)
Private Const PREVIEW_FONT   As String = "Consolas"                     ' fixed-pitch -> runs stay aligned
Private Const PREVIEW_SIZE   As Single = 9
Private Const PREVIEW_X0     As Single = 2                              ' left margin inside Frame_RulePreview
Private Const PREVIEW_Y0     As Single = 2                              ' top margin inside Frame_RulePreview
' Deterministic fixed-pitch cell (points). Each run label is sized EXPLICITLY from these rather than from
' per-label AutoSize, which does not re-fit a runtime Label in a Frame after its caption is set (that was
' the "@Ce*]=" truncation bug). Slightly generous so a run never clips; width = Len*CHARW + INSET. Tunable.
Private Const PREVIEW_CHARW  As Single = PREVIEW_SIZE * 0.55            ' FALLBACK advance/char if MeasureCharAdvance fails (Consolas ~0.55 em); tunable
Private Const PREVIEW_INSET  As Single = 4                             ' a Label's fixed left/right text inset
Private Const PREVIEW_CHARH  As Single = PREVIEW_SIZE * 1.7            ' line height (incl. margins)

' ============================================================
' MASTER SWITCH - CheckBox -> ARES_Auto_Properties
' ============================================================

Private Sub Main_CheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    ' Enter toggles the checkbox too (uniform with buttons; Space already toggles natively).
    If Shift = 0 And KeyCode = vbKeyReturn Then Main_CheckBox.value = Not Main_CheckBox.value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.Main_CheckBox_KeyUp"
End Sub

Private Sub Main_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(Main_CheckBox.value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_AUTO_PROPERTIES.value <> sVal Then
        SetLocked True
        ARESConfig.ARES_AUTO_PROPERTIES.value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.Main_CheckBox_Change"
End Sub

' ============================================================
' CUSTOM PROPERTY LIST - Edit button + hidden TextBox -> ARES_Custom_Property_List
' ============================================================

Private Sub Edit_PropertyList_Command_Click()
    On Error GoTo ErrorHandler
    If Not mbLocked Then
        SetLocked True
        TextBox_PropertyList.value = ARESConfig.ARES_CUSTOM_PROPERTY_LIST.value
        TextBox_PropertyList.Visible = True
        Edit_PropertyList_Command.Visible = False
        TextBox_PropertyList.SetFocus
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.Edit_PropertyList_Command_Click"
End Sub

Private Sub TextBox_PropertyList_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    FormUXHelper.CommitInlineEdit TextBox_PropertyList, Edit_PropertyList_Command, ARESConfig.ARES_CUSTOM_PROPERTY_LIST
    SetLocked False
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.TextBox_PropertyList_Exit"
End Sub

Private Sub TextBox_PropertyList_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    FormUXHelper.NoteInlineKeyDown KeyCode, Shift
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.TextBox_PropertyList_KeyDown"
End Sub

Private Sub TextBox_PropertyList_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            TextBox_PropertyList_Exit returnB
            Edit_PropertyList_Command.SetFocus
        Case FormUXKeyCancel
            FormUXHelper.RevertInlineEdit TextBox_PropertyList, ARESConfig.ARES_CUSTOM_PROPERTY_LIST
            TextBox_PropertyList_Exit returnB
            Edit_PropertyList_Command.SetFocus
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.TextBox_PropertyList_KeyUp"
End Sub

' ============================================================
' PROPERTY RULES - editable ComboBox -> ARES_Property_Rules (one rule at a time)
' The rules string is split on ";" into individual entries. Pick a rule to load it for editing;
' committing (Enter or focus-out) replaces the picked rule, appends a free-typed one, or removes an
' emptied one. Every commit is validated (PropertyTagging.ValidateAndNormalizeRule) - a malformed rule is
' refused with a one-shot status (status only, never logged) and the combo re-seeds to the stored value (no write).
' ============================================================

' Re-seed the ComboBox from ARES_Property_Rules (split on ";", trimmed, empties dropped) and clear the
' edit area + the tracked edit index. Called on init, after every accepted commit, and on revert. The
' LAST item is an empty "new rule" sentinel (UI only): picking it returns to add mode; it is NEVER
' written to ARES_Property_Rules (CommitRuleEdit maps a sentinel pick to -1 and skips empty items).
Private Sub SeedRulesCombo()
    On Error GoTo ErrorHandler
    ComboBox_Rules.Clear
    Dim vRules As Variant, i As Long
    vRules = Split(ARESConfig.ARES_PROPERTY_RULES.value, RULE_SEP)
    For i = LBound(vRules) To UBound(vRules)
        If Len(Trim(vRules(i))) > 0 Then ComboBox_Rules.AddItem Trim(vRules(i))
    Next i
    ComboBox_Rules.AddItem ""                    ' trailing "new rule" sentinel (UI only, never stored)
    ComboBox_Rules.text = ""
    mRuleEditIndex = -1
    RenderRulePreview                            ' refresh the coloured preview (usually clears - text is "")
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.SeedRulesCombo"
End Sub

' Capture the picked list index BEFORE the edited text diverges from it. Only a clean pick
' (Text = List(ListIndex)) sets mRuleEditIndex; once the user edits, ListIndex may reset to -1 but the
' captured index survives. Nested Ifs (VBA has no short-circuit; List(ListIndex) must not run at -1).
Private Sub ComboBox_Rules_Change()
    On Error GoTo ErrorHandler
    If mbLocked Then Exit Sub
    If ComboBox_Rules.ListIndex >= 0 Then
        If ComboBox_Rules.text = ComboBox_Rules.List(ComboBox_Rules.ListIndex) Then
            ' The last item is the empty "new rule" sentinel -> ADD mode (-1), NOT its list index. Without
            ' this, an empty-text commit from the sentinel would hit remove semantics with an index past
            ' the real rules. (Nested Ifs, never And.)
            If ComboBox_Rules.ListIndex = ComboBox_Rules.ListCount - 1 Then
                mRuleEditIndex = -1
            Else
                mRuleEditIndex = ComboBox_Rules.ListIndex
            End If
        End If
    End If
    RenderRulePreview                            ' live coloured preview of the edited text (read-only)
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.ComboBox_Rules_Change"
End Sub

Private Sub ComboBox_Rules_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    FormUXHelper.NoteInlineKeyDown KeyCode, Shift
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.ComboBox_Rules_KeyDown"
End Sub

Private Sub ComboBox_Rules_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            CommitRuleEdit
        Case FormUXKeyCancel
            SeedRulesCombo                      ' revert: drop the edit, no write
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.ComboBox_Rules_KeyUp"
End Sub

Private Sub ComboBox_Rules_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    CommitRuleEdit                              ' focus-out commit (same path as Enter)
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.ComboBox_Rules_Exit"
End Sub

' Apply the current edit to the rules list and write it back. Uses mRuleEditIndex to decide:
'   index >= 0, text non-empty -> REPLACE that rule
'   index >= 0, text empty      -> REMOVE that rule
'   index <  0, text non-empty -> APPEND a new rule (free-typed, or picked from the "new rule" sentinel)
'   index <  0, text empty      -> no-op (covers a sentinel pick with no text - nothing is written)
' A non-empty edited rule is validated first; a refusal shows PropertyRuleInvalid (status only - a
' mistyped rule is expected user input, not a fault, so it is never logged) and re-seeds to the stored
' value (no partial write). On accept: reassemble with ";", write, RefreshRules,
' re-seed. The empty "new rule" sentinel is skipped in the rebuild so it is never stored. Both Enter
' (KeyUp) and focus-out (Exit) route here - one commit path.
Private Sub CommitRuleEdit()
    On Error GoTo ErrorHandler
    If mbLocked Then Exit Sub                   ' re-entrance guard (a commit already running)

    Dim sEdited As String
    Dim bHasIndex As Boolean
    sEdited = Trim(ComboBox_Rules.text)
    bHasIndex = (mRuleEditIndex >= 0)

    ' Free-typed nothing: no change - just re-seed a clean combo. (Nested Ifs, never And.)
    If Not bHasIndex Then
        If Len(sEdited) = 0 Then
            SeedRulesCombo
            Exit Sub
        End If
    End If

    ' Validate + normalise the single edited rule (an empty text is a delete and needs no validation).
    ' On success sCanonical holds the CANONICAL stored form (what gets written); on failure a status is
    ' shown and nothing is written. The reason itself is discarded (status-only - a mistyped rule is
    ' expected user input, not a fault, so it is never logged).
    Dim sCanonical As String
    sCanonical = ""
    If Len(sEdited) > 0 Then
        Dim sReason As String
        sReason = PropertyTagging.ValidateAndNormalizeRule(sEdited, sCanonical)
        If Len(sReason) > 0 Then
            LangManager.ShowStatusT "PropertyRuleInvalid"
            SeedRulesCombo                       ' revert to the last-good list
            Exit Sub
        End If
    End If

    ' Rebuild the rules list from the current combo items with the edit applied. Empty items (only ever
    ' the "new rule" sentinel) are skipped so the sentinel is never written to ARES_Property_Rules.
    Dim rebuilt() As String
    Dim nCount As Long, i As Long, w As Long
    Dim bIsTarget As Boolean
    Dim sItem As String
    nCount = ComboBox_Rules.ListCount
    ReDim rebuilt(0 To nCount)                   ' room for every item + one possible append
    w = 0
    For i = 0 To nCount - 1
        bIsTarget = False
        If bHasIndex Then
            If i = mRuleEditIndex Then bIsTarget = True
        End If
        If bIsTarget Then
            If Len(sEdited) > 0 Then
                rebuilt(w) = sCanonical          ' replace with the canonical form
                w = w + 1
            End If
            ' empty -> skip (remove)
        Else
            sItem = Trim(ComboBox_Rules.List(i))
            If Len(sItem) > 0 Then               ' skip the empty "new rule" sentinel (never stored)
                rebuilt(w) = sItem
                w = w + 1
            End If
        End If
    Next i
    If Not bHasIndex Then                        ' free-typed rule -> append (canonical form)
        If Len(sEdited) > 0 Then
            rebuilt(w) = sCanonical
            w = w + 1
        End If
    End If

    Dim sJoined As String
    If w = 0 Then
        sJoined = ""
    Else
        ReDim Preserve rebuilt(0 To w - 1)
        sJoined = Join(rebuilt, RULE_SEP)
    End If

    SetLocked True
    ARESConfig.ARES_PROPERTY_RULES.value = sJoined
    PropertyTagging.RefreshRules                 ' apply the edited rules live, no restart
    SeedRulesCombo
    SetLocked False
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.CommitRuleEdit"
End Sub

' ============================================================
' COLOURED SYNTAX PREVIEW (read-only) -> Frame_RulePreview
' Re-renders the text currently in ComboBox_Rules as runtime Labels in a fixed-pitch font: @ pink, & dark
' blue, ! orange, */? green, contradiction segments RED (via PropertyTagging.RuleHasNoEffect), the rest
' monochrome. Renders the CANONICAL form when the rule is valid (so the preview shows what will be stored
' and the red segments map onto exact substrings), the raw text when invalid (metachar colours only).
' Purely visual: no config write, not in the tab chain, fail-safe (any render fault clears the preview).
' ============================================================

' Rebuild the preview for the text currently in ComboBox_Rules. Fail-safe: any error clears the preview
' and logs - a broken cosmetic preview must never break the rule editor it decorates.
Private Sub RenderRulePreview()
    On Error GoTo ErrorHandler

    ClearPreview

    Dim sText As String
    sText = ComboBox_Rules.text
    If Len(Trim(sText)) = 0 Then Exit Sub        ' empty -> cleared (no Labels)

    ' Validate + normalise (read-only, 13-2): render the canonical form when valid, the raw text when not.
    Dim sCanonical As String
    Dim sReason As String
    Dim sRender As String
    Dim bValid As Boolean
    sReason = PropertyTagging.ValidateAndNormalizeRule(sText, sCanonical)
    bValid = (Len(sReason) = 0)
    If bValid Then
        sRender = sCanonical
    Else
        sRender = sText
    End If
    If Len(sRender) = 0 Then Exit Sub

    ' Per-character colour + bold maps. Colours: metachars, then overlay contradiction red (valid rules
    ' only - RuleHasNoEffect is defined only on a syntactically valid rule). Bold: the keyword tokens
    ' Lvl/Cell/Type (only) render bold. MarkSegmentsRed touches colours ONLY, so bold never shifts the red
    ' character positions.
    Dim colours() As Long
    Dim bolds() As Boolean
    colours = BuildColourMap(sRender)
    bolds = BuildBoldMap(sRender)
    If bValid Then
        Dim segs() As String
        If PropertyTagging.RuleHasNoEffect(sCanonical, segs) Then
            MarkSegmentsRed sCanonical, segs, colours
        End If
    End If

    ' Stop laying runs out once past the Frame's inner width - a long rule is clipped on the right (no wrap).
    Dim dMaxX As Single
    dMaxX = 0
    On Error Resume Next
    dMaxX = Frame_RulePreview.InsideWidth
    On Error GoTo ErrorHandler

    ' Coalesce neighbours sharing the SAME (colour, bold) into runs, laid out left-to-right. Each run
    ' advances x by its EXACT text width so runs sit seamlessly (fixed pitch; transparent labels hide the
    ' generous right padding that overlaps the next run).
    Dim n As Long, i As Long, runStart As Long
    Dim x As Single
    Dim bBreak As Boolean
    Dim bUnderline As Boolean
    Dim dCharW As Single
    n = Len(sRender)
    x = PREVIEW_X0
    runStart = 1
    bUnderline = Not bValid
    dCharW = MeasureCharAdvance()                 ' real rendered advance (GDI pixel/DPI), not the theoretical const

    ' Invalid rule: an error cue - a red-bold ballot-X (U+2717) marker + space at the head, and every text
    ' run underlined (spell-checker style). The glyph is built with ChrW (NEVER a literal - the .frm is ANSI
    ' and would eat it); if MSForms/Consolas does not render it, replace ChrW(&H2717) with "X". No red
    ' analysis runs on an invalid rule, so the marker's leading x offset affects no colour/segment mapping.
    If Not bValid Then
        x = EmitRun(ChrW(&H2717) & " ", PREVIEW_RED, True, False, dCharW, x)
    End If

    For i = 1 To n
        bBreak = False
        If i = n Then
            bBreak = True
        Else
            If colours(i + 1) <> colours(runStart) Then bBreak = True
            If bolds(i + 1) <> bolds(runStart) Then bBreak = True
        End If
        If bBreak Then
            x = EmitRun(Mid(sRender, runStart, i - runStart + 1), colours(runStart), bolds(runStart), bUnderline, dCharW, x)
            runStart = i + 1
        End If
        If dMaxX > 0 Then
            If x >= dMaxX Then Exit For
        End If
    Next i
    Exit Sub

ErrorHandler:
    ClearPreview
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.RenderRulePreview"
End Sub

' Per-character colour map (1-based) for the render string: metachars coloured, everything else monochrome.
' Colouring metachars wherever they appear (not only at bracket depth 0) is acceptable for a preview - a
' literal & / @ inside [...] is rare and harmless. Contradiction red is overlaid separately (valid only).
Private Function BuildColourMap(ByVal s As String) As Long()
    On Error GoTo ErrorHandler

    Dim colours() As Long
    Dim n As Long, i As Long
    Dim ch As String
    n = Len(s)
    If n < 1 Then n = 1                           ' never ReDim(1 To 0)
    ReDim colours(1 To n)
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        Select Case ch
            Case "@"
                colours(i) = PREVIEW_PINK
            Case "&"
                colours(i) = PREVIEW_BLUE
            Case "!"
                colours(i) = PREVIEW_ORANGE
            Case "*", "?"
                colours(i) = PREVIEW_GREEN
            Case Else
                colours(i) = PREVIEW_MONO
        End Select
    Next i
    BuildColourMap = colours
    Exit Function

ErrorHandler:
    ' Fail-safe: an all-monochrome map of the right size so the caller can still render.
    ReDim colours(1 To IIf(Len(s) < 1, 1, Len(s)))
    For i = LBound(colours) To UBound(colours)
        colours(i) = PREVIEW_MONO
    Next i
    BuildColourMap = colours
End Function

' Overlay PREVIEW_RED on the characters of each conflicting segment (an exact substring of the canonical
' text, returned by RuleHasNoEffect). Search from a running position so repeated segments map to distinct
' ranges; red overrides any metachar colour in range (a dead condition reads as unmistakably red).
Private Sub MarkSegmentsRed(ByVal sCanonical As String, ByRef segments() As String, ByRef colours() As Long)
    On Error GoTo ErrorHandler

    Dim si As Long, p As Long, k As Long, pos As Long
    pos = 1
    For si = LBound(segments) To UBound(segments)
        If Len(segments(si)) > 0 Then
            p = InStr(pos, sCanonical, segments(si))
            If p > 0 Then
                For k = p To p + Len(segments(si)) - 1
                    If k >= LBound(colours) Then
                        If k <= UBound(colours) Then colours(k) = PREVIEW_RED
                    End If
                Next k
                pos = p + Len(segments(si))
            End If
        End If
    Next si
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.MarkSegmentsRed"
End Sub

' Per-character BOLD map (1-based): the keyword tokens Lvl / Cell / Type (and only them) render bold - not
' the names inside [...] and not the metacharacters. A keyword is the run of letters immediately before a
' bracket-depth-0 "[".
Private Function BuildBoldMap(ByVal s As String) As Boolean()
    On Error GoTo ErrorHandler

    Dim bold() As Boolean
    Dim n As Long, i As Long, depth As Long, ch As String
    n = Len(s)
    If n < 1 Then n = 1
    ReDim bold(1 To n)                           ' defaults all False
    depth = 0
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch = "[" Then
            If depth = 0 Then MarkKeywordBold s, i - 1, bold
            depth = depth + 1
        ElseIf ch = "]" Then
            depth = depth - 1
        End If
    Next i
    BuildBoldMap = bold
    Exit Function

ErrorHandler:
    ' Fail-safe: an all-regular map of the right size.
    ReDim bold(1 To IIf(Len(s) < 1, 1, Len(s)))
    BuildBoldMap = bold
End Function

' If the letters ending at endPos form a Lvl / Cell / Type keyword, mark that character range bold.
Private Sub MarkKeywordBold(ByVal s As String, ByVal endPos As Long, ByRef bold() As Boolean)
    On Error GoTo ErrorHandler

    If endPos < 1 Then Exit Sub

    Dim startPos As Long, j As Long
    startPos = endPos + 1
    For j = endPos To 1 Step -1
        If IsLetter(Mid(s, j, 1)) Then
            startPos = j
        Else
            Exit For
        End If
    Next j
    If startPos > endPos Then Exit Sub           ' no letters immediately before the "["

    Dim k As Long
    Select Case UCase(Mid(s, startPos, endPos - startPos + 1))
        Case "LVL", "CELL", "TYPE"
            For k = startPos To endPos
                If k >= LBound(bold) Then
                    If k <= UBound(bold) Then bold(k) = True
                End If
            Next k
    End Select
    Exit Sub

ErrorHandler:
End Sub

' True when ch is an ASCII letter A-Z / a-z (nested Ifs, no And; module is Option Compare Binary).
Private Function IsLetter(ByVal ch As String) As Boolean
    IsLetter = False
    If Len(ch) = 0 Then Exit Function
    Dim u As String
    u = UCase(ch)
    If u >= "A" Then
        If u <= "Z" Then IsLetter = True
    End If
End Function

' Create one Label for a coloured (optionally bold) run, size it EXPLICITLY (deterministic fixed-pitch
' cell, no AutoSize), and return the x for the NEXT run advanced by the run's EXACT text width so runs sit
' seamlessly (the generous right padding overlaps the next run but is invisible - transparent BackStyle).
Private Function EmitRun(ByVal sRun As String, ByVal lColour As Long, ByVal bBold As Boolean, ByVal bUnderline As Boolean, ByVal dCharW As Single, ByVal x As Single) As Single
    On Error GoTo ErrorHandler

    EmitRun = x
    If Len(sRun) = 0 Then Exit Function

    Dim oLbl As MSForms.Label
    Set oLbl = AddPreviewLabel()
    If oLbl Is Nothing Then Exit Function

    ' Order: font (name/size/bold/underline) FIRST, then caption, then EXPLICIT width/height (AutoSize off,
    ' no wrap) so the label never depends on a per-label AutoSize recalculation. dCharW is the RUNTIME-measured
    ' advance (MeasureCharAdvance) - Width = Len*dCharW + INSET is generous so the text never clips; x below
    ' advances by the EXACT Len*dCharW so successive runs join seamlessly with no cumulative pixel drift.
    ' Consolas keeps the same advance in bold/underlined, so a regular-weight calibration suffices.
    oLbl.Visible = False
    ' A fresh per-label StdFont: MSForms controls added via Controls.Add SHARE their container's Font object,
    ' so mutating oLbl.Font.* would contaminate every other (and future) label - a valid rule would then
    ' inherit the bold/underline left by a previous invalid render. Assigning a new StdFont isolates this
    ' label's font. (No oLbl.Font.* mutation anywhere after this.)
    Dim f As stdole.StdFont
    Set f = New stdole.StdFont
    f.Name = PREVIEW_FONT
    f.Size = PREVIEW_SIZE
    f.bold = bBold
    f.Underline = bUnderline
    Set oLbl.Font = f
    oLbl.AutoSize = False
    oLbl.WordWrap = False
    oLbl.Caption = sRun
    oLbl.ForeColor = lColour
    oLbl.Width = Len(sRun) * dCharW + PREVIEW_INSET
    oLbl.Height = PREVIEW_CHARH
    oLbl.Left = x
    oLbl.Top = PREVIEW_Y0
    oLbl.Visible = True
    EmitRun = x + Len(sRun) * dCharW
    Exit Function

ErrorHandler:
    ' Silent fail-safe (per-run): a failed run just does not advance x. RenderRulePreview is the single
    ' logger of the render path; per-run logging would spam if Controls.Add is systemically unavailable.
    EmitRun = x
End Function

' Measure the REAL rendered character advance once per render: a hidden calibration label (from the pool,
' never shown) gets a fresh StdFont (regular weight - Consolas' bold advance is identical) and a known
' 64-char etalon, then AutoSize = True is set AFTER the caption (toggled off->on to force the recompute -
' the inverse of the round-1 order bug). dCharW = .Width / N absorbs the fixed label margin over N=64
' (negligible). SANITY CHECK: an implausible result (AutoSize did not recompute) falls back to the
' theoretical PREVIEW_CHARW constant (no worse than before). This removes the GDI pixel/DPI drift a
' theoretical advance caused (progressive position creep + last-glyph crop).
Private Function MeasureCharAdvance() As Single
    On Error GoTo ErrorHandler

    MeasureCharAdvance = PREVIEW_CHARW            ' fallback = theoretical constant

    Const CAL_N As Long = 64
    Dim oCal As MSForms.Label
    Set oCal = AddPreviewLabel()                 ' hidden, tracked in the pool (cleared next render - no leak)
    If oCal Is Nothing Then Exit Function

    Dim f As stdole.StdFont
    Set f = New stdole.StdFont
    f.Name = PREVIEW_FONT
    f.Size = PREVIEW_SIZE
    f.bold = False
    f.Underline = False
    Set oCal.Font = f

    oCal.WordWrap = False
    oCal.Caption = String(CAL_N, "M")            ' known N-char etalon
    oCal.AutoSize = False
    oCal.AutoSize = True                          ' AFTER the caption (round-1 order inverted) -> forces the recompute

    If oCal.Width > 0 Then
        Dim dMeasured As Single
        dMeasured = oCal.Width / CAL_N
        ' Plausible fixed-pitch advance? else AutoSize did not recompute -> keep the fallback constant.
        If dMeasured >= PREVIEW_SIZE * 0.3 Then
            If dMeasured <= PREVIEW_SIZE * 1.2 Then
                MeasureCharAdvance = dMeasured
            End If
        End If
    End If
    Exit Function

ErrorHandler:
    MeasureCharAdvance = PREVIEW_CHARW
End Function

' Create a runtime Label inside Frame_RulePreview, track its name for removal, return it (Nothing on fault).
Private Function AddPreviewLabel() As MSForms.Label
    On Error GoTo ErrorHandler

    EnsurePreviewCollection
    mlPreviewSeq = mlPreviewSeq + 1
    Dim sName As String
    sName = "lblPreview" & mlPreviewSeq

    Dim oLbl As MSForms.Label
    Set oLbl = Frame_RulePreview.Controls.Add("Forms.Label.1", sName, False)   ' created hidden; shown once sized
    moPreviewLabels.Add sName
    oLbl.BackStyle = fmBackStyleTransparent
    Set AddPreviewLabel = oLbl
    Exit Function

ErrorHandler:
    ' Silent fail-safe: a failed Controls.Add just yields no Label (the run is skipped). See EmitRun.
    Set AddPreviewLabel = Nothing
End Function

' Remove exactly the runtime preview Labels created by the previous render (bounded control count, no leak).
' Silent (On Error Resume Next): a stale name is skipped; cleanup faults must never surface.
Private Sub ClearPreview()
    On Error Resume Next
    If Not moPreviewLabels Is Nothing Then
        Dim v As Variant
        For Each v In moPreviewLabels
            Frame_RulePreview.Controls.Remove CStr(v)
        Next v
    End If
    Set moPreviewLabels = New Collection
    On Error GoTo 0
End Sub

' Lazily create the preview-label name collection.
Private Sub EnsurePreviewCollection()
    If moPreviewLabels Is Nothing Then Set moPreviewLabels = New Collection
End Sub

' ============================================================
' FORM LIFECYCLE
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Set moPreviewLabels = New Collection         ' before SeedControls -> SeedRulesCombo -> RenderRulePreview
    Me.Caption = GetTranslation("PropertyTaggingGUIOptionsCaption")
    ' Checkbox caption lives on the checkbox: Tab-focus visible + the text toggles the box
    Main_CheckBox.Caption = GetTranslation("PropertyTaggingGUIOptionsMain_LabelCaption")
    Edit_PropertyList_Command.Caption = GetTranslation("PropertyTaggingGUIOptionsEditList_CommandCaption")

    ' Tooltips
    FormUXHelper.SetTip Main_CheckBox, "PropertyTaggingGUIOptionsMain_LabelTip"
    FormUXHelper.SetTip Edit_PropertyList_Command, "PropertyTaggingGUIOptionsEditList_CommandTip"
    FormUXHelper.SetTip TextBox_PropertyList, "PropertyTaggingGUIOptionsEditList_CommandTip"
    FormUXHelper.SetTip ComboBox_Rules, "PropertyTaggingGUIOptionsEditRules_CommandTip"

    ' Match ComboBox_Rules' font to the coloured preview (fixed-pitch Consolas), so the combo text and the
    ' preview below it share the same character widths - otherwise the columns don't line up (proportional
    ' Tahoma vs fixed Consolas). Assign a FRESH StdFont via Set (never ComboBox_Rules.Font.* - the shared
    ' Font-object trap from round 3). This also switches the dropdown list to Consolas (intended, coherent).
    ' The 9pt font can be a touch taller than a combo drawn for 8pt, so bump the height at runtime ONLY if it
    ' would be too short (conditional -> no overlap with the preview Frame when the designer sized it right;
    ' Asketyll can instead set the combo height in the designer).
    Dim fCombo As stdole.StdFont
    Set fCombo = New stdole.StdFont
    fCombo.Name = PREVIEW_FONT
    fCombo.Size = PREVIEW_SIZE
    Set ComboBox_Rules.Font = fCombo
    If ComboBox_Rules.Height < PREVIEW_SIZE * 2 Then ComboBox_Rules.Height = PREVIEW_SIZE * 2

    ' Restore-defaults button
    Reset_Command.Caption = GetTranslation("FormResetDefaultsCaption")
    FormUXHelper.SetTip Reset_Command, "FormResetDefaultsTip"

    SeedControls
    FormPlacement.RestoreFormPosition Me, Me.Name
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.UserForm_Initialize"
End Sub

' Re-seed all controls from the current config values.
Private Sub SeedControls()
    On Error GoTo ErrorHandler
    If ARESConfig.ARES_AUTO_PROPERTIES.value Then
        Main_CheckBox.value = "True"
    Else
        Main_CheckBox.value = "False"
    End If
    TextBox_PropertyList.Visible = False
    SeedRulesCombo
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.SeedControls"
End Sub

' Restore every option this form edits to its default value, persist, then re-seed.
Private Sub Reset_Command_Click()
    On Error GoTo ErrorHandler
    FormUXHelper.PersistDefault ARESConfig.ARES_AUTO_PROPERTIES
    FormUXHelper.PersistDefault ARESConfig.ARES_CUSTOM_PROPERTY_LIST
    FormUXHelper.PersistDefault ARESConfig.ARES_PROPERTY_RULES
    PropertyTagging.RefreshRules
    SeedControls
    LangManager.ShowStatusT "FormDefaultsRestored"
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.Reset_Command_Click"
End Sub

' Any error path must call SetLocked False so controls are never left disabled.
Private Sub SetLocked(ByVal bState As Boolean)
    On Error GoTo ErrorHandler
    mbLocked = bState
    FormUXHelper.SetControlsLocked Me, bState
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.SetLocked"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHandler

    If mbLocked Then
        Cancel = True
        If TextBox_PropertyList.Visible Then FormUXHelper.NudgeActiveEdit TextBox_PropertyList
    Else
        ' The ComboBox is the sole rules editor, so flush a pending combo edit on click-X (MSForms does
        ' not guarantee the combo's _Exit fires on teardown). CommitRuleEdit is re-entrance-guarded and
        ' idempotent: a valid edit is written, an invalid one is dropped with PropertyRuleInvalid, and an
        ' already-committed / empty state is a harmless no-op (RA7). No partial write on any path.
        CommitRuleEdit
        FormPlacement.SaveFormPosition Me, Me.Name
        command.OnPropertyTaggingGUIClosed
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.UserForm_QueryClose"
End Sub



