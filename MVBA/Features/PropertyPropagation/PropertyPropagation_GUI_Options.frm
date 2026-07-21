VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PropertyPropagation_GUI_Options 
   Caption         =   "PropertyPropagation_GUI_Options"
   ClientHeight    =   1935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5295
   OleObjectBlob   =   "PropertyPropagation_GUI_Options.frx":0000
End
Attribute VB_Name = "PropertyPropagation_GUI_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm: PropertyPropagation_GUI_Options
' Description: Options panel for Property Propagation - master switch (ARES_Property_Propagation),
'              trigger cell-name | -list (ARES_Propagation_Cells), and target custom property
'              (ARES_Propagation_Property). Mirrors PropertyTagging_GUI_Options (checkbox + inline-
'              edit reveal) plus a Null-safe property ComboBox with conditional enable, borrowed from
'              ExportLengthInReg_GUI_Options.
'
'              DESIGNER (manual, Asketyll) - controls required with EXACTLY these names:
'                Main_CheckBox (CheckBox), Edit_CellNames_Command (CommandButton),
'                TextBox_CellNames (TextBox, hidden until reveal), Property_Label (Label),
'                ComboBox_Property (ComboBox, Style = 2 dropdown list), Reset_Command (CommandButton).
'              StartUpPosition = 0 Manual. Tab order: master -> cell-names -> property -> reset.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass, ARESConfigClass, CustomPropertyHandler, FormUXHelper, FormPlacement, Command
Option Explicit

Private mbLocked As Boolean

' ============================================================
' MASTER SWITCH - CheckBox -> ARES_Property_Propagation
' ============================================================

Private Sub Main_CheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    ' Enter toggles the checkbox too (uniform with buttons; Space already toggles natively).
    If Shift = 0 And KeyCode = vbKeyReturn Then Main_CheckBox.value = Not Main_CheckBox.value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.Main_CheckBox_KeyUp"
End Sub

Private Sub Main_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(Main_CheckBox.value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_PROPERTY_PROPAGATION.value <> sVal Then
        SetLocked True
        ARESConfig.ARES_PROPERTY_PROPAGATION.value = sVal
        SetLocked False
    End If
    ApplyMasterEnabled
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.Main_CheckBox_Change"
End Sub

' ============================================================
' TRIGGER CELL NAMES - Edit button + hidden TextBox -> ARES_Propagation_Cells (| -list)
' ============================================================

Private Sub Edit_CellNames_Command_Click()
    On Error GoTo ErrorHandler
    If Not mbLocked Then
        SetLocked True
        TextBox_CellNames.value = ARESConfig.ARES_PROPAGATION_CELLS.value
        TextBox_CellNames.Visible = True
        Edit_CellNames_Command.Visible = False
        TextBox_CellNames.SetFocus
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.Edit_CellNames_Command_Click"
End Sub

Private Sub TextBox_CellNames_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    FormUXHelper.CommitInlineEdit TextBox_CellNames, Edit_CellNames_Command, ARESConfig.ARES_PROPAGATION_CELLS
    SetLocked False
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.TextBox_CellNames_Exit"
End Sub

Private Sub TextBox_CellNames_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    FormUXHelper.NoteInlineKeyDown KeyCode, Shift
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.TextBox_CellNames_KeyDown"
End Sub

Private Sub TextBox_CellNames_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            TextBox_CellNames_Exit returnB
            Edit_CellNames_Command.SetFocus
        Case FormUXKeyCancel
            FormUXHelper.RevertInlineEdit TextBox_CellNames, ARESConfig.ARES_PROPAGATION_CELLS
            TextBox_CellNames_Exit returnB
            Edit_CellNames_Command.SetFocus
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.TextBox_CellNames_KeyUp"
End Sub

' ============================================================
' TARGET PROPERTY - ComboBox (populated from ARES_Custom_Property_List) -> ARES_Propagation_Property
' Enabled only while the master switch is ON.
' ============================================================

Private Sub ComboBox_Property_Change()
    On Error GoTo ErrorHandler
    If mbLocked Then Exit Sub
    Dim sVal As String
    ' Null-safe read: a dropdown-list combo with no selection returns Null (assigning it to a String
    ' would raise Error 94). Nested If (not And) per the no-short-circuit cheatsheet rule.
    If IsNull(ComboBox_Property.value) Then
        sVal = ""
    Else
        sVal = Trim(CStr(ComboBox_Property.value))
    End If
    If ARESConfig.ARES_PROPAGATION_PROPERTY.value <> sVal Then
        SetLocked True
        ARESConfig.ARES_PROPAGATION_PROPERTY.value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.ComboBox_Property_Change"
End Sub

' ============================================================
' FORM LIFECYCLE
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption = GetTranslation("PropagationGUIOptionsCaption")
    ' Checkbox caption lives on the checkbox: Tab-focus visible + the text toggles the box
    Main_CheckBox.Caption = GetTranslation("PropagationGUIOptionsMain_LabelCaption")
    Edit_CellNames_Command.Caption = GetTranslation("PropagationGUIOptionsCells_LabelCaption")
    Property_Label.Caption = GetTranslation("PropagationGUIOptionsProperty_LabelCaption")

    ' Tooltips
    FormUXHelper.SetTip Main_CheckBox, "PropagationGUIOptionsMain_LabelTip"
    FormUXHelper.SetTip Edit_CellNames_Command, "PropagationGUIOptionsCells_LabelTip"
    FormUXHelper.SetTip TextBox_CellNames, "PropagationGUIOptionsCells_LabelTip"
    FormUXHelper.SetTip Property_Label, "PropagationGUIOptionsProperty_LabelTip"
    FormUXHelper.SetTip ComboBox_Property, "PropagationGUIOptionsProperty_LabelTip"

    ' Restore-defaults button
    Reset_Command.Caption = GetTranslation("FormResetDefaultsCaption")
    FormUXHelper.SetTip Reset_Command, "FormResetDefaultsTip"

    SeedControls
    FormPlacement.RestoreFormPosition Me, Me.Name
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.UserForm_Initialize"
End Sub

' Re-seed all controls from the current config values.
Private Sub SeedControls()
    On Error GoTo ErrorHandler

    Main_CheckBox.value = (UCase(Trim(ARESConfig.ARES_PROPERTY_PROPAGATION.value)) = "TRUE")
    TextBox_CellNames.Visible = False

    ' Property combo: populate from the managed custom-property names, then seed the current value
    ' Null-safe (a dropdown-list combo rejects an out-of-list value) - set .Value only when the stored
    ' name is a list member, else .ListIndex = -1 (M1).
    ComboBox_Property.Clear
    Dim propNames()  As String
    Dim pi           As Long
    Dim sSel         As String
    Dim bFound       As Boolean
    sSel = Trim(ARESConfig.ARES_PROPAGATION_PROPERTY.value)
    bFound = False
    propNames = CustomPropertyHandler.GetCustomPropertyNames()
    For pi = LBound(propNames) To UBound(propNames)
        If Len(Trim(propNames(pi))) > 0 Then
            ComboBox_Property.AddItem propNames(pi)
            If StrComp(Trim(propNames(pi)), sSel, vbTextCompare) = 0 Then bFound = True
        End If
    Next pi
    If bFound Then
        ComboBox_Property.value = sSel
    Else
        ComboBox_Property.ListIndex = -1
    End If

    ApplyMasterEnabled
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.SeedControls"
End Sub

' Restore every option this form edits to its default value, persist, then re-seed.
Private Sub Reset_Command_Click()
    On Error GoTo ErrorHandler
    FormUXHelper.PersistDefault ARESConfig.ARES_PROPERTY_PROPAGATION
    FormUXHelper.PersistDefault ARESConfig.ARES_PROPAGATION_CELLS
    FormUXHelper.PersistDefault ARESConfig.ARES_PROPAGATION_PROPERTY
    SeedControls
    LangManager.ShowStatusT "FormDefaultsRestored"
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.Reset_Command_Click"
End Sub

' Conditional grey-out: the property combo + its label are meaningful only while the master switch is
' ON. Re-asserted after every SetControlsLocked(False) and on every SeedControls. Runs only on
' unlock/seed, when the combo cannot hold focus, so disabling it never ejects focus.
Private Sub ApplyMasterEnabled()
    On Error GoTo ErrorHandler
    ' Null-safe read: a triple-state checkbox could yield Null, which errors on a direct Boolean
    ' assign; "If ... = True" treats Null as False (bOn defaults False).
    Dim bOn As Boolean
    If Main_CheckBox.value = True Then bOn = True
    ComboBox_Property.Enabled = bOn
    Property_Label.Enabled = bOn
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.ApplyMasterEnabled"
End Sub

' Any error path must call SetLocked False so controls are never left disabled.
Private Sub SetLocked(ByVal bState As Boolean)
    On Error GoTo ErrorHandler
    mbLocked = bState
    FormUXHelper.SetControlsLocked Me, bState
    ' SetControlsLocked(False) re-enables every combo/checkbox, so re-assert the conditional
    ' property grey-out after each global unlock.
    If Not bState Then ApplyMasterEnabled
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.SetLocked"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHandler

    If mbLocked Then
        Cancel = True
        If TextBox_CellNames.Visible Then FormUXHelper.NudgeActiveEdit TextBox_CellNames
    Else
        FormPlacement.SaveFormPosition Me, Me.Name
        command.OnPropagationGUIClosed
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.UserForm_QueryClose"
End Sub

