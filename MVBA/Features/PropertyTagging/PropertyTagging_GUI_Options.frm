VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PropertyTagging_GUI_Options 
   Caption         =   "UserForm1"
   ClientHeight    =   2655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3735
   OleObjectBlob   =   "PropertyTagging_GUI_Options.frx":0000
   StartUpPosition =   0  'Manual
End
Attribute VB_Name = "PropertyTagging_GUI_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm: PropertyTagging_GUI_Options
' Description: UserForm for editing custom-property (Property Tagging) options.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass, ARESConfigClass, PropertyTagging
Option Explicit

Private mbLocked As Boolean

' ============================================================
' MASTER SWITCH - CheckBox -> ARES_Auto_Properties
' ============================================================

Private Sub Main_CheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    ' Enter toggles the checkbox too (uniform with buttons; Space already toggles natively).
    If Shift = 0 And KeyCode = vbKeyReturn Then Main_CheckBox.Value = Not Main_CheckBox.Value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.Main_CheckBox_KeyUp"
End Sub

Private Sub Main_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(Main_CheckBox.Value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_AUTO_PROPERTIES.Value <> sVal Then
        SetLocked True
        ARESConfig.ARES_AUTO_PROPERTIES.Value = sVal
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
        TextBox_PropertyList.Value = ARESConfig.ARES_CUSTOM_PROPERTY_LIST.Value
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
' PROPERTY RULES - Edit button + hidden TextBox -> ARES_Property_Rules
' On write, refresh PropertyTagging's parsed cache so the new rules take effect immediately.
' ============================================================

Private Sub Edit_Rules_Command_Click()
    On Error GoTo ErrorHandler
    If Not mbLocked Then
        SetLocked True
        TextBox_Rules.Value = ARESConfig.ARES_PROPERTY_RULES.Value
        TextBox_Rules.Visible = True
        Edit_Rules_Command.Visible = False
        TextBox_Rules.SetFocus
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.Edit_Rules_Command_Click"
End Sub

Private Sub TextBox_Rules_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    If FormUXHelper.CommitInlineEdit(TextBox_Rules, Edit_Rules_Command, ARESConfig.ARES_PROPERTY_RULES) Then
        PropertyTagging.RefreshRules            ' apply the edited rules live, no restart
    End If
    SetLocked False
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.TextBox_Rules_Exit"
End Sub

Private Sub TextBox_Rules_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    FormUXHelper.NoteInlineKeyDown KeyCode, Shift
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.TextBox_Rules_KeyDown"
End Sub

Private Sub TextBox_Rules_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            TextBox_Rules_Exit returnB
            Edit_Rules_Command.SetFocus
        Case FormUXKeyCancel
            FormUXHelper.RevertInlineEdit TextBox_Rules, ARESConfig.ARES_PROPERTY_RULES
            TextBox_Rules_Exit returnB
            Edit_Rules_Command.SetFocus
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.TextBox_Rules_KeyUp"
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
    Edit_Rules_Command.Caption = GetTranslation("PropertyTaggingGUIOptionsEditRules_CommandCaption")

    ' Tooltips
    FormUXHelper.SetTip Main_CheckBox, "PropertyTaggingGUIOptionsMain_LabelTip"
    FormUXHelper.SetTip Edit_PropertyList_Command, "PropertyTaggingGUIOptionsEditList_CommandTip"
    FormUXHelper.SetTip TextBox_PropertyList, "PropertyTaggingGUIOptionsEditList_CommandTip"
    FormUXHelper.SetTip Edit_Rules_Command, "PropertyTaggingGUIOptionsEditRules_CommandTip"
    FormUXHelper.SetTip TextBox_Rules, "PropertyTaggingGUIOptionsEditRules_CommandTip"


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
    If ARESConfig.ARES_AUTO_PROPERTIES.Value Then
        Main_CheckBox.Value = "True"
    Else
        Main_CheckBox.Value = "False"
    End If
    TextBox_PropertyList.Visible = False
    TextBox_Rules.Visible = False
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
        Select Case True
            Case TextBox_PropertyList.Visible
                FormUXHelper.NudgeActiveEdit TextBox_PropertyList
            Case TextBox_Rules.Visible
                FormUXHelper.NudgeActiveEdit TextBox_Rules
        End Select
    Else
        FormPlacement.SaveFormPosition Me, Me.Name
        command.OnPropertyTaggingGUIClosed
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.UserForm_QueryClose"
End Sub

