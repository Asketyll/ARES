VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PropertyCalculation_GUI_Options 
   Caption         =   "PropertyPropagation_GUI_Options"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5550
   OleObjectBlob   =   "PropertyCalculation_GUI_Options.frx":0000
End
Attribute VB_Name = "PropertyCalculation_GUI_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm: PropertyCalculation_GUI_Options
' Description: Options panel for Property Calculation - the value-calc master switch
'              (ARES_Property_Calc), the detach-empty option (ARES_Calc_Detach_Empty), and the calc rules
'              (ARES_Calc_Rules, epic 14). Calc rules are edited one at a time through an editable ComboBox
'              (split on ";"): pick a rule -> edit -> commit replaces it; free-type + commit appends; empty
'              + commit removes it. Every commit is validated by PropertyCalculation.ValidateAndNormalizeCalcRule,
'              so a malformed rule is refused (status-only, never logged) instead of saved.
'
'              The editable-ComboBox mechanics and the read-only COLOURED SYNTAX PREVIEW below
'              ComboBox_CalcRules are provided by the shared RuleEditorUX module (the same code the tag rule
'              editor uses - no divergence). This form stays a thin consumer: it runs the calc grammar
'              (ValidateAndNormalizeCalcRule + CalcRuleHasNoEffect) and hands RuleEditorUX the render string
'              + validity + red segments + the bold-keyword list (the target/condition keywords Prop/Lvl/
'              Cell/Type AND every calc source keyword - the list itself is the kw() array in
'              RenderCurrentCalcPreview; grammar reference: calc-rules-grammar.md).
'
'              DESIGNER (manual) - controls required with EXACTLY these names:
'                Main_CheckBox (CheckBox, value master), DetachEmpty_CheckBox (CheckBox, detach-empty
'                option; caption set in code), ComboBox_CalcRules (ComboBox, Style = 0 fmStyleDropDownCombo
'                EDITABLE - the sole per-rule editor), Frame_CalcPreview (Frame, render surface directly
'                BELOW ComboBox_CalcRules - the coloured Labels are created at runtime, NONE in the
'                designer), Reset_Command (CommandButton), Help_Command (CommandButton, opens the wiki page
'                - placed beside Reset_Command). Help_Command's Picture (a "?" icon) and MousePointer (14 =
'                fmMousePointerHelp) are set HERE IN THE DESIGNER, not in code - a Win32 icon-load from code
'                was rejected as too risky for a cosmetic icon (see the comment beside
'                FormUXHelper.SetTip Help_Command in UserForm_Initialize).
'              StartUpPosition = 0 Manual. Tab order: master -> detach-empty -> calc-rules-combo -> reset ->
'              help (Frame_CalcPreview is a non-focusable container, not in the tab order).
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass, ARESConfigClass, PropertyCalculation, RuleEditorUX, FormUXHelper, FormPlacement, Command
Option Explicit

Private mbLocked As Boolean

' The ComboBox list index the user picked before editing its text (-1 = new / free-typed). Maintained via
' RuleEditorUX.CaptureEditIndex on every _Change (a clean pick sets it; typing preserves it).
Private mCalcRuleEditIndex As Long

' ============================================================
' MASTER SWITCH - CheckBox -> ARES_Property_Calc
' ============================================================

Private Sub Main_CheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    ' Enter toggles the checkbox too (uniform with buttons; Space already toggles natively).
    If Shift = 0 And KeyCode = vbKeyReturn Then Main_CheckBox.value = Not Main_CheckBox.value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_GUI_Options.Main_CheckBox_KeyUp"
End Sub

Private Sub Main_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(Main_CheckBox.value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_PROPERTY_CALC.value <> sVal Then
        SetLocked True
        ARESConfig.ARES_PROPERTY_CALC.value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_GUI_Options.Main_CheckBox_Change"
End Sub

' ============================================================
' DETACH-EMPTY OPTION - CheckBox -> ARES_Calc_Detach_Empty (round-4)
' When on, an emptied value is DETACHED (via the tagger) instead of cleared. Independent of the master
' switch (it may be on while the master is off).
' ============================================================

Private Sub DetachEmpty_CheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    ' Enter toggles the checkbox too (uniform with buttons; Space already toggles natively).
    If Shift = 0 And KeyCode = vbKeyReturn Then DetachEmpty_CheckBox.value = Not DetachEmpty_CheckBox.value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_GUI_Options.DetachEmpty_CheckBox_KeyUp"
End Sub

Private Sub DetachEmpty_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(DetachEmpty_CheckBox.value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_CALC_DETACH_EMPTY.value <> sVal Then
        SetLocked True
        ARESConfig.ARES_CALC_DETACH_EMPTY.value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_GUI_Options.DetachEmpty_CheckBox_Change"
End Sub

' ============================================================
' CALC RULES - editable ComboBox -> ARES_Calc_Rules (one rule at a time)
' The editable-ComboBox mechanics + coloured preview are provided by RuleEditorUX (the SAME code as the tag
' rule editor); this form runs the calc grammar (PropertyCalculation.ValidateAndNormalizeCalcRule /
' CalcRuleHasNoEffect) and delegates the rest.
' ============================================================

' Re-seed the ComboBox from ARES_Calc_Rules and refresh the preview. The seed (clear/split/sentinel/
' reset-text) is RuleEditorUX's; the edit index and the preview render stay here.
Private Sub SeedCalcRulesCombo()
    On Error GoTo ErrorHandler
    RuleEditorUX.SeedRulesCombo ComboBox_CalcRules, ARESConfig.ARES_CALC_RULES.value
    mCalcRuleEditIndex = -1
    RenderCurrentCalcPreview
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_GUI_Options.SeedCalcRulesCombo"
End Sub

' On every combo change: update the tracked edit index (RuleEditorUX.CaptureEditIndex preserves it while
' typing, sets it on a clean pick) and re-render the coloured preview of the edited text.
Private Sub ComboBox_CalcRules_Change()
    On Error GoTo ErrorHandler
    If mbLocked Then Exit Sub
    mCalcRuleEditIndex = RuleEditorUX.CaptureEditIndex(ComboBox_CalcRules, mCalcRuleEditIndex)
    RenderCurrentCalcPreview
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_GUI_Options.ComboBox_CalcRules_Change"
End Sub

Private Sub ComboBox_CalcRules_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    FormUXHelper.NoteInlineKeyDown KeyCode, Shift
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_GUI_Options.ComboBox_CalcRules_KeyDown"
End Sub

Private Sub ComboBox_CalcRules_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            CommitCalcRuleEdit
        Case FormUXKeyCancel
            SeedCalcRulesCombo                  ' revert: drop the edit, no write
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_GUI_Options.ComboBox_CalcRules_KeyUp"
End Sub

Private Sub ComboBox_CalcRules_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    CommitCalcRuleEdit                          ' focus-out commit (same path as Enter)
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_GUI_Options.ComboBox_CalcRules_Exit"
End Sub

' Apply the current edit to the calc-rules list and write it back. This form keeps the GRAMMAR-SPECIFIC
' lines: read the edited text, validate + normalise via PropertyCalculation (reason -> status, LEAVE the
' typed text as-is, no write; canonical otherwise), then delegate the rebuild to RuleEditorUX.RebuildRules,
' write ARES_Calc_Rules, RefreshCalcRules, re-seed. A refusal shows CalcRuleInvalid (status only - never
' logged) and does NOT re-seed the combo: wiping the box back to "" on an invalid edit reads as "my rule got
' deleted" (it wasn't - nothing is written on a refusal) when it was really just the user's typo bounced back
' blank with no way to fix it in place. Both Enter (KeyUp) and focus-out (Exit) route here - one commit path.
Private Sub CommitCalcRuleEdit()
    On Error GoTo ErrorHandler
    If mbLocked Then Exit Sub                   ' re-entrance guard (a commit already running)

    Dim sEdited As String
    Dim bHasIndex As Boolean
    sEdited = Trim(ComboBox_CalcRules.text)
    bHasIndex = (mCalcRuleEditIndex >= 0)

    ' Free-typed nothing: no change - just re-seed a clean combo. (Nested Ifs, never And.)
    If Not bHasIndex Then
        If Len(sEdited) = 0 Then
            SeedCalcRulesCombo
            Exit Sub
        End If
    End If

    ' Validate + normalise the single edited rule (an empty text is a delete and needs no validation).
    ' On success sCanonical holds the CANONICAL stored form; on failure a status is shown and nothing is
    ' written. The reason itself is discarded (status-only - a mistyped rule is expected input, not a fault).
    Dim sCanonical As String
    sCanonical = ""
    If Len(sEdited) > 0 Then
        Dim sReason As String
        sReason = PropertyCalculation.ValidateAndNormalizeCalcRule(sEdited, sCanonical)
        If Len(sReason) > 0 Then
            LangManager.ShowStatusT "CalcRuleInvalid"
            RenderCurrentCalcPreview             ' re-assert the invalid marker; the typed text stays put so the user can fix it
            Exit Sub
        End If
    End If

    ' Rebuild the ";"-joined value (replace/append/remove, sentinel excluded) - delegated to RuleEditorUX.
    Dim sJoined As String
    sJoined = RuleEditorUX.RebuildRules(ComboBox_CalcRules, mCalcRuleEditIndex, sCanonical)

    SetLocked True
    ARESConfig.ARES_CALC_RULES.value = sJoined
    PropertyCalculation.RefreshCalcRules         ' apply the edited calc rules live, no restart
    SeedCalcRulesCombo
    SetLocked False
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_GUI_Options.CommitCalcRuleEdit"
End Sub

' Compute this grammar's preview data for the text currently in ComboBox_CalcRules and hand it to
' RuleEditorUX: validate + normalise -> canonical (valid) or raw (invalid); when valid, CalcRuleHasNoEffect
' -> the red segments; kw() below is the bold-keyword list - target/condition keywords plus every calc
' source, argument-less ones included (RuleEditorUX bolds a bare keyword too). It renders the coloured runs
' read-only. Keep kw() in step with ParseSource's Case list when a source is added.
Private Sub RenderCurrentCalcPreview()
    On Error GoTo ErrorHandler

    Dim sText As String, sCanonical As String, sReason As String, sRender As String
    Dim bValid As Boolean
    Dim segs() As String
    Dim kw(23) As String

    ReDim segs(0 To 0)
    segs(0) = ""
    sRender = ""

    sText = ComboBox_CalcRules.text
    If Len(Trim(sText)) > 0 Then
        sReason = PropertyCalculation.ValidateAndNormalizeCalcRule(sText, sCanonical)
        bValid = (Len(sReason) = 0)
        If bValid Then
            sRender = sCanonical
            PropertyCalculation.CalcRuleHasNoEffect sCanonical, segs   ' fills segs with the contradiction segments
        Else
            sRender = sText
        End If
    End If

    kw(0) = "Prop"
    kw(1) = "Lvl"
    kw(2) = "Cell"
    kw(3) = "Type"
    kw(4) = "CellText"
    kw(5) = "Value"
    kw(6) = "Coord"
    kw(7) = "Id"
    kw(8) = "CellCoord"
    kw(9) = "CellId"
    kw(10) = "CellLvl"
    kw(11) = "Color"
    kw(12) = "CellColor"
    kw(13) = "Style"
    kw(14) = "CellStyle"
    kw(15) = "Weight"
    kw(16) = "CellWeight"
    kw(17) = "Length"
    kw(18) = "GroupLength"
    kw(19) = "LvlColor"
    kw(20) = "LvlStyle"
    kw(21) = "LvlWeight"
    kw(22) = "GroupColor"
    kw(23) = "GroupProp"
    RuleEditorUX.RenderPreview Frame_CalcPreview, sRender, bValid, segs, kw
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_GUI_Options.RenderCurrentCalcPreview"
End Sub

' ============================================================
' FORM LIFECYCLE
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption = GetTranslation("CalculationGUIOptionsCaption")
    ' Checkbox captions live on the checkboxes: Tab-focus visible + the text toggles the box
    Main_CheckBox.Caption = GetTranslation("CalculationGUIOptionsMain_LabelCaption")
    DetachEmpty_CheckBox.Caption = GetTranslation("CalculationGUIOptionsDetachEmpty_LabelCaption")

    ' Tooltips
    FormUXHelper.SetTip Main_CheckBox, "CalculationGUIOptionsMain_LabelTip"
    FormUXHelper.SetTip DetachEmpty_CheckBox, "CalculationGUIOptionsDetachEmpty_LabelTip"
    FormUXHelper.SetTip ComboBox_CalcRules, "CalculationGUIOptionsCalcRules_Tip"

    ' Match ComboBox_CalcRules' font to the coloured preview (fixed-pitch), so the combo text and the preview
    ' below it line up (delegated to RuleEditorUX - fresh StdFont, conditional height bump).
    RuleEditorUX.MatchComboFont ComboBox_CalcRules

    ' Restore-defaults button
    Reset_Command.Caption = GetTranslation("FormResetDefaultsCaption")
    FormUXHelper.SetTip Reset_Command, "FormResetDefaultsTip"

    ' Help button: the ComboBox tooltip has no room for the full calc-rules grammar reference - this opens
    ' the wiki page instead. Its Picture (a "?" icon) is set in the DESIGNER, not here - owned the same way
    ' as MousePointer/tab order (a Win32 API icon load was considered and rejected: the PICTDESC struct
    ' layout differs 32/64-bit and a mistake there is an unrecoverable access violation, not a catchable
    ' VBA error - not worth the risk for a cosmetic icon that the designer sets safely in one click).
    FormUXHelper.SetTip Help_Command, "FormHelpTip"

    SeedControls
    FormPlacement.RestoreFormPosition Me, Me.Name
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_GUI_Options.UserForm_Initialize"
End Sub

' Re-seed the two checkboxes + the calc-rules combo from the current config values.
Private Sub SeedControls()
    On Error GoTo ErrorHandler

    Main_CheckBox.value = (UCase(Trim(ARESConfig.ARES_PROPERTY_CALC.value)) = "TRUE")
    ' Detach-empty option is independent of the master switch (seeded like Main_CheckBox).
    DetachEmpty_CheckBox.value = (UCase(Trim(ARESConfig.ARES_CALC_DETACH_EMPTY.value)) = "TRUE")
    SeedCalcRulesCombo
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_GUI_Options.SeedControls"
End Sub

' Open the wiki page with the full calc-rules syntax reference (EN/FR resolved by ARES_Language) - the
' ComboBox tooltip alone cannot hold the full grammar + all sixteen source keywords legibly.
Private Sub Help_Command_Click()
    On Error GoTo ErrorHandler
    command.OpenARESWikiPage "Property-Calculation", "Calcul-de-Propriete"
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_GUI_Options.Help_Command_Click"
End Sub

' Restore every option this form edits to its default value, persist, then re-seed.
Private Sub Reset_Command_Click()
    On Error GoTo ErrorHandler
    If Not FormUXHelper.ConfirmReset() Then Exit Sub
    FormUXHelper.PersistDefault ARESConfig.ARES_PROPERTY_CALC
    FormUXHelper.PersistDefault ARESConfig.ARES_CALC_DETACH_EMPTY
    FormUXHelper.PersistDefault ARESConfig.ARES_CALC_RULES
    PropertyCalculation.RefreshCalcRules
    SeedControls
    LangManager.ShowStatusT "FormDefaultsRestored"
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_GUI_Options.Reset_Command_Click"
End Sub

' Any error path must call SetLocked False so controls are never left disabled.
Private Sub SetLocked(ByVal bState As Boolean)
    On Error GoTo ErrorHandler
    mbLocked = bState
    FormUXHelper.SetControlsLocked Me, bState
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_GUI_Options.SetLocked"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHandler

    If mbLocked Then
        Cancel = True
    Else
        ' The ComboBox is the sole calc-rules editor, so flush a pending combo edit on click-X (MSForms does
        ' not guarantee the combo's _Exit fires on teardown). CommitCalcRuleEdit is re-entrance-guarded and
        ' idempotent: a valid edit is written, an invalid one is dropped with CalcRuleInvalid, and an
        ' already-committed / empty state is a harmless no-op (RA7). No partial write on any path.
        CommitCalcRuleEdit
        FormPlacement.SaveFormPosition Me, Me.Name
        command.OnPropertyCalculationGUIClosed
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculation_GUI_Options.UserForm_QueryClose"
End Sub


