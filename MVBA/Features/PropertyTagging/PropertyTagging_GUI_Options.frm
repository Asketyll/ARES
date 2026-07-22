VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PropertyTagging_GUI_Options 
   Caption         =   "PropertyTagging_GUI_Options"
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3735
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
'              PropertyTagging.ValidateRuleSyntax, so a malformed rule (chiefly the "|"-instead-of-";"
'              mistake) is refused instead of saved.
'
'              DESIGNER (manual, Asketyll) - controls required with EXACTLY these names:
'                Main_CheckBox (CheckBox, master), Edit_PropertyList_Command (CommandButton) +
'                TextBox_PropertyList (TextBox, hidden reveal - property list), ComboBox_Rules (ComboBox,
'                Style = 0 fmStyleDropDownCombo EDITABLE - the sole per-rule editor),
'                Reset_Command (CommandButton).
'              StartUpPosition = 0 Manual. Tab order: master -> property-list -> rules-combo -> reset.
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
' emptied one. Every commit is validated (PropertyTagging.ValidateRuleSyntax) - a malformed rule is
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

    ' Validate the single edited rule (an empty text is a delete and needs no validation).
    If Len(sEdited) > 0 Then
        Dim sReason As String
        sReason = PropertyTagging.ValidateRuleSyntax(sEdited)
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
                rebuilt(w) = sEdited             ' replace
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
    If Not bHasIndex Then                        ' free-typed rule -> append
        If Len(sEdited) > 0 Then
            rebuilt(w) = sEdited
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


