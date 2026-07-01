VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PropertyTagging_GUI_Options 
   Caption         =   "UserForm1"
   ClientHeight    =   1575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "PropertyTagging_GUI_Options.frx":0000
   StartUpPosition =   1  'CenterOwner
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

Private Sub TextBox_PropertyList_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            TextBox_PropertyList_Exit returnB
        Case FormUXKeyCancel
            FormUXHelper.RevertInlineEdit TextBox_PropertyList, ARESConfig.ARES_CUSTOM_PROPERTY_LIST
            TextBox_PropertyList_Exit returnB
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

Private Sub TextBox_Rules_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            TextBox_Rules_Exit returnB
        Case FormUXKeyCancel
            FormUXHelper.RevertInlineEdit TextBox_Rules, ARESConfig.ARES_PROPERTY_RULES
            TextBox_Rules_Exit returnB
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
    Main_Label.Caption = GetTranslation("PropertyTaggingGUIOptionsMain_LabelCaption")
    Edit_PropertyList_Command.Caption = GetTranslation("PropertyTaggingGUIOptionsEditList_CommandCaption")
    Edit_Rules_Command.Caption = GetTranslation("PropertyTaggingGUIOptionsEditRules_CommandCaption")

    ' Tooltips (AC-6)
    FormUXHelper.SetTip Main_CheckBox, "PropertyTaggingGUIOptionsMain_LabelTip"
    FormUXHelper.SetTip Main_Label, "PropertyTaggingGUIOptionsMain_LabelTip"
    FormUXHelper.SetTip Edit_PropertyList_Command, "PropertyTaggingGUIOptionsEditList_CommandTip"
    FormUXHelper.SetTip TextBox_PropertyList, "PropertyTaggingGUIOptionsEditList_CommandTip"
    FormUXHelper.SetTip Edit_Rules_Command, "PropertyTaggingGUIOptionsEditRules_CommandTip"
    FormUXHelper.SetTip TextBox_Rules, "PropertyTaggingGUIOptionsEditRules_CommandTip"

    ' Keyboard order + mnemonics (AC-7) - existing controls only
    Main_CheckBox.TabIndex = 0
    Edit_PropertyList_Command.TabIndex = 1
    Edit_Rules_Command.TabIndex = 2
    Main_CheckBox.Accelerator = "A"
    Edit_PropertyList_Command.Accelerator = "P"
    Edit_Rules_Command.Accelerator = "R"

    If ARESConfig.ARES_AUTO_PROPERTIES.Value Then
        Main_CheckBox.Value = "True"
    Else
        Main_CheckBox.Value = "False"
    End If

    TextBox_PropertyList.Visible = False
    TextBox_Rules.Visible = False
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.UserForm_Initialize"
End Sub

' Explicit-state lock (AC-2/AC-8): replaces the toggle Locked()/CheckControlForLock pair.
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
        command.OnPropertyTaggingGUIClosed
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.UserForm_QueryClose"
End Sub

