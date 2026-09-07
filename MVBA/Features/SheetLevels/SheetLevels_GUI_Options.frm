VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SheetLevels_GUI_Options 
   Caption         =   "SheetLevels_GUI_Options"
   ClientHeight    =   1455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2775
   OleObjectBlob   =   "SheetLevels_GUI_Options.frx":0000
   StartUpPosition =   0  'Manual
End
Attribute VB_Name = "SheetLevels_GUI_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm: SheetLevels_GUI_Options
' Description: UserForm for editing Sheet Levels options (the sheet-model name pattern, and whether the
'              sheets' references are processed too).
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass, ARESConfigClass, FormUXHelper, FormPlacement
Option Explicit

Private mbLocked As Boolean

' ============================================================
' MODEL NAME PATTERN - Edit button + hidden TextBox
' ============================================================

Private Sub Edit_Name_Command_Click()
    On Error GoTo ErrorHandler
    If Not mbLocked Then
        SetLocked True
        TextBox_Name.value = ARESConfig.ARES_SHEET_LEVELS_MODEL_NAME.value
        TextBox_Name.Visible = True
        Edit_Name_Command.Visible = False
        TextBox_Name.SetFocus
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels_GUI_Options.Edit_Name_Command_Click"
End Sub

Private Sub TextBox_Name_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    FormUXHelper.CommitInlineEdit TextBox_Name, Edit_Name_Command, ARESConfig.ARES_SHEET_LEVELS_MODEL_NAME
    Edit_Name_Command.Caption = GetTranslation("SheetLevelsGUIOptionsEditName_CommandCaption", ARESConfig.ARES_SHEET_LEVELS_MODEL_NAME.value)
    SetLocked False
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels_GUI_Options.TextBox_Name_Exit"
End Sub

Private Sub TextBox_Name_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    FormUXHelper.NoteInlineKeyDown KeyCode, Shift
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels_GUI_Options.TextBox_Name_KeyDown"
End Sub

Private Sub TextBox_Name_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            TextBox_Name_Exit returnB
            Edit_Name_Command.SetFocus
        Case FormUXKeyCancel
            FormUXHelper.RevertInlineEdit TextBox_Name, ARESConfig.ARES_SHEET_LEVELS_MODEL_NAME
            TextBox_Name_Exit returnB
            Edit_Name_Command.SetFocus
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels_GUI_Options.TextBox_Name_KeyUp"
End Sub

' ============================================================
' PROCESS REFERENCES - CheckBox (surfaces ARES_Sheet_Levels_Attachments)
' On (the default), a matching sheet's references get their levels turned on too - which is where a
' folio's drawing usually comes from.
' ============================================================

Private Sub Attachments_CheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    ' Enter toggles the checkbox too (uniform with buttons; Space already toggles natively).
    If Shift = 0 And KeyCode = vbKeyReturn Then Attachments_CheckBox.value = Not Attachments_CheckBox.value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels_GUI_Options.Attachments_CheckBox_KeyUp"
End Sub

Private Sub Attachments_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(Attachments_CheckBox.value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_SHEET_LEVELS_ATTACHMENTS.value <> sVal Then
        SetLocked True
        ARESConfig.ARES_SHEET_LEVELS_ATTACHMENTS.value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels_GUI_Options.Attachments_CheckBox_Change"
End Sub

' ============================================================
' FORM LIFECYCLE
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption = GetTranslation("SheetLevelsGUIOptionsCaption")
    Edit_Name_Command.WordWrap = True
    Attachments_CheckBox.Caption = GetTranslation("SheetLevelsGUIOptionsAttachments_LabelCaption")

    ' Tooltips
    FormUXHelper.SetTip Edit_Name_Command, "SheetLevelsGUIOptionsEditName_CommandTip"
    FormUXHelper.SetTip Attachments_CheckBox, "SheetLevelsGUIOptionsAttachments_LabelTip"

    ' Restore-defaults button
    Reset_Command.Caption = GetTranslation("FormResetDefaultsCaption")
    FormUXHelper.SetTip Reset_Command, "FormResetDefaultsTip"

    SeedControls
    FormPlacement.RestoreFormPosition Me, Me.name
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels_GUI_Options.UserForm_Initialize"
End Sub

' Re-seed all controls from the current config values.
Private Sub SeedControls()
    On Error GoTo ErrorHandler
    Edit_Name_Command.Caption = GetTranslation("SheetLevelsGUIOptionsEditName_CommandCaption", ARESConfig.ARES_SHEET_LEVELS_MODEL_NAME.value)

    ' Anything that is not literally "False" is True - the same reading the engine applies, so the box
    ' can never show the opposite of what a run will do.
    Attachments_CheckBox.value = (UCase(Trim(ARESConfig.ARES_SHEET_LEVELS_ATTACHMENTS.value)) <> "FALSE")

    TextBox_Name.Visible = False
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels_GUI_Options.SeedControls"
End Sub

' Restore every option this form edits to its default value, persist, then re-seed.
Private Sub Reset_Command_Click()
    On Error GoTo ErrorHandler
    If Not FormUXHelper.ConfirmReset() Then Exit Sub
    FormUXHelper.PersistDefault ARESConfig.ARES_SHEET_LEVELS_MODEL_NAME
    FormUXHelper.PersistDefault ARESConfig.ARES_SHEET_LEVELS_ATTACHMENTS
    SeedControls
    LangManager.ShowStatusT "FormDefaultsRestored"
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels_GUI_Options.Reset_Command_Click"
End Sub

' Any error path must call SetLocked False so controls are never left disabled.
Private Sub SetLocked(ByVal bState As Boolean)
    On Error GoTo ErrorHandler
    mbLocked = bState
    FormUXHelper.SetControlsLocked Me, bState
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels_GUI_Options.SetLocked"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHandler

    If mbLocked Then
        Cancel = True
        If TextBox_Name.Visible Then FormUXHelper.NudgeActiveEdit TextBox_Name
    Else
        FormPlacement.SaveFormPosition Me, Me.name
        command.OnSheetLevelsGUIClosed
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "SheetLevels_GUI_Options.UserForm_QueryClose"
End Sub
