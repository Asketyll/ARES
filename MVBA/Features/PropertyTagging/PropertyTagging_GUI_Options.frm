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
'              The editable-ComboBox mechanics and the read-only COLOURED SYNTAX PREVIEW below
'              ComboBox_Rules are provided by the shared RuleEditorUX module (epic 14 - the Calculation
'              form reuses the same code). This form stays a thin consumer: it runs its grammar
'              (ValidateAndNormalizeRule + RuleHasNoEffect) and hands RuleEditorUX the render string +
'              validity + red segments + bold-keyword list {Lvl,Cell,Type}; RuleEditorUX renders the runs.
'
'              DESIGNER (manual, Asketyll) - controls required with EXACTLY these names:
'                Main_CheckBox (CheckBox, master), Edit_PropertyList_Command (CommandButton) +
'                TextBox_PropertyList (TextBox, hidden reveal - property list), ComboBox_Rules (ComboBox,
'                Style = 0 fmStyleDropDownCombo EDITABLE - the sole per-rule editor),
'                Frame_RulePreview (Frame, render surface directly BELOW ComboBox_Rules for the runtime
'                coloured preview - the coloured Labels are created at runtime, NONE in the designer),
'                Reset_Command (CommandButton).
'              StartUpPosition = 0 Manual. Tab order: master -> property-list -> rules-combo -> reset
'              (Frame_RulePreview is a non-focusable container, not in the tab order).
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass, ARESConfigClass, PropertyTagging, RuleEditorUX, FormUXHelper, FormPlacement, CustomPropertyHandler, Command
Option Explicit

Private mbLocked As Boolean

' The ComboBox list index the user picked before editing its text (-1 = new / free-typed). Maintained via
' RuleEditorUX.CaptureEditIndex on every _Change (a clean pick sets it; typing preserves it).
Private mRuleEditIndex As Long

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
    ' The list decides which ItemTypes ARES manages, so a real change refreshes MicroStation's Item Type
    ' state. CommitInlineEdit returns True ONLY when it actually wrote, which is exactly the "once per
    ' real commit" seam: Enter/Esc call this sub manually and the ensuing focus change can fire the real
    ' _Exit a second time, but that pass sees box = stored value and returns False; Esc reverts before
    ' committing, so a cancel returns False too. Sent directly rather than deferred to idle - a form event
    ' is user-driven, with MicroStation between operations, unlike the DGN-open path.
    If FormUXHelper.CommitInlineEdit(TextBox_PropertyList, Edit_PropertyList_Command, ARESConfig.ARES_CUSTOM_PROPERTY_LIST) Then
        CustomPropertyHandler.RefreshItemTypes
    End If
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
' The editable-ComboBox mechanics + coloured preview are provided by RuleEditorUX; this form runs the
' grammar (PropertyTagging.ValidateAndNormalizeRule / RuleHasNoEffect) and delegates the rest.
' ============================================================

' Re-seed the ComboBox from ARES_Property_Rules and refresh the preview. The seed (clear/split/sentinel/
' reset-text) is RuleEditorUX's; the edit index and the preview render stay here.
Private Sub SeedRulesCombo()
    On Error GoTo ErrorHandler
    RuleEditorUX.SeedRulesCombo ComboBox_Rules, ARESConfig.ARES_PROPERTY_RULES.value
    mRuleEditIndex = -1
    RenderCurrentPreview
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.SeedRulesCombo"
End Sub

' On every combo change: update the tracked edit index (RuleEditorUX.CaptureEditIndex preserves it while
' typing, sets it on a clean pick) and re-render the coloured preview of the edited text.
Private Sub ComboBox_Rules_Change()
    On Error GoTo ErrorHandler
    If mbLocked Then Exit Sub
    mRuleEditIndex = RuleEditorUX.CaptureEditIndex(ComboBox_Rules, mRuleEditIndex)
    RenderCurrentPreview
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

' Apply the current edit to the rules list and write it back. This form keeps the GRAMMAR-SPECIFIC lines:
' read the edited text, validate + normalise via PropertyTagging (reason -> status + re-seed, no write;
' canonical otherwise), then delegate the rebuild to RuleEditorUX.RebuildRules, write ARES_Property_Rules,
' RefreshRules, re-seed. A refusal shows PropertyRuleInvalid (status only - never logged). Both Enter
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
    ' On success sCanonical holds the CANONICAL stored form; on failure a status is shown and nothing is
    ' written. The reason itself is discarded (status-only - a mistyped rule is expected input, not a fault).
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

    ' Rebuild the ";"-joined value (replace/append/remove, sentinel excluded) - delegated to RuleEditorUX.
    Dim sJoined As String
    sJoined = RuleEditorUX.RebuildRules(ComboBox_Rules, mRuleEditIndex, sCanonical)

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

' Compute this grammar's preview data for the text currently in ComboBox_Rules and hand it to RuleEditorUX:
' validate + normalise -> canonical (valid) or raw (invalid); when valid, RuleHasNoEffect -> the red
' segments; the tag bold-keywords are {Lvl,Cell,Type}. RuleEditorUX renders the coloured runs (read-only).
Private Sub RenderCurrentPreview()
    On Error GoTo ErrorHandler

    Dim sText As String, sCanonical As String, sReason As String, sRender As String
    Dim bValid As Boolean
    Dim segs() As String
    Dim kw(2) As String

    ReDim segs(0 To 0)
    segs(0) = ""
    sRender = ""

    sText = ComboBox_Rules.text
    If Len(Trim(sText)) > 0 Then
        sReason = PropertyTagging.ValidateAndNormalizeRule(sText, sCanonical)
        bValid = (Len(sReason) = 0)
        If bValid Then
            sRender = sCanonical
            PropertyTagging.RuleHasNoEffect sCanonical, segs   ' fills segs with the contradiction segments
        Else
            sRender = sText
        End If
    End If

    kw(0) = "Lvl"
    kw(1) = "Cell"
    kw(2) = "Type"
    RuleEditorUX.RenderPreview Frame_RulePreview, sRender, bValid, segs, kw
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.RenderCurrentPreview"
End Sub

' ============================================================
' FORM LIFECYCLE
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption = GetTranslation("PropertyTaggingGUIOptionsCaption")
    ' Checkbox caption lives on the checkbox: Tab-focus visible + the text toggles the box
    Main_CheckBox.Caption = GetTranslation("PropertyTaggingGUIOptionsMain_LabelCaption")
    Edit_PropertyList_Command.Caption = GetTranslation("PropertyTaggingGUIOptionsEditList_CommandCaption")

    ' Tooltips
    FormUXHelper.SetTip Main_CheckBox, "PropertyTaggingGUIOptionsMain_LabelTip"
    FormUXHelper.SetTip Edit_PropertyList_Command, "PropertyTaggingGUIOptionsEditList_CommandTip"
    FormUXHelper.SetTip TextBox_PropertyList, "PropertyTaggingGUIOptionsEditList_CommandTip"
    FormUXHelper.SetTip ComboBox_Rules, "PropertyTaggingGUIOptionsEditRules_CommandTip"

    ' Match ComboBox_Rules' font to the coloured preview (fixed-pitch), so the combo text and the preview
    ' below it line up (delegated to RuleEditorUX - fresh StdFont, conditional height bump).
    RuleEditorUX.MatchComboFont ComboBox_Rules

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



