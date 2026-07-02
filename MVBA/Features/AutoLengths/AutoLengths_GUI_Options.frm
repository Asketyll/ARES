VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AutoLengths_GUI_Options 
   Caption         =   "Edit auto lengths options:"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3015
   OleObjectBlob   =   "AutoLengths_GUI_Options.frx":0000
   StartUpPosition =   0  'Manual
End
Attribute VB_Name = "AutoLengths_GUI_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm: AutoLengths_GUI_Options
' Description: This UserForm is used for editing the option of AutoLengths
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass, ARESConfigClass, ARESConstants, FormUXHelper
Option Explicit

Private mbLocked As Boolean

' ============================================================
' CHECKBOXES -> config booleans
' ============================================================

Private Sub Main_CheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    ' Enter toggles the checkbox too (uniform with buttons; Space already toggles natively).
    If Shift = 0 And KeyCode = vbKeyReturn Then Main_CheckBox.Value = Not Main_CheckBox.Value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.Main_CheckBox_KeyUp"
End Sub

Private Sub Main_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(Main_CheckBox.Value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_AUTO_LENGTHS.Value <> sVal Then
        SetLocked True
        ARESConfig.ARES_AUTO_LENGTHS.Value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.Main_CheckBox_Change"
End Sub

Private Sub Color_CheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    If Shift = 0 And KeyCode = vbKeyReturn Then Color_CheckBox.Value = Not Color_CheckBox.Value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.Color_CheckBox_KeyUp"
End Sub

Private Sub Color_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(Color_CheckBox.Value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_UPDATE_COLOR_WITH_LENGTH.Value <> sVal Then
        SetLocked True
        ARESConfig.ARES_UPDATE_COLOR_WITH_LENGTH.Value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.Color_CheckBox_Change"
End Sub

Private Sub Only_Color_CheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    If Shift = 0 And KeyCode = vbKeyReturn Then Only_Color_CheckBox.Value = Not Only_Color_CheckBox.Value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.Only_Color_CheckBox_KeyUp"
End Sub

Private Sub Only_Color_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(Only_Color_CheckBox.Value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_ONLY_COLOR.Value <> sVal Then
        SetLocked True
        ARESConfig.ARES_ONLY_COLOR.Value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.Only_Color_CheckBox_Change"
End Sub

Private Sub Cell_CheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    If Shift = 0 And KeyCode = vbKeyReturn Then Cell_CheckBox.Value = Not Cell_CheckBox.Value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.Cell_CheckBox_KeyUp"
End Sub

Private Sub Cell_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(Cell_CheckBox.Value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_UPDATE_ATLASCELLLABEL.Value <> sVal Then
        SetLocked True
        ARESConfig.ARES_UPDATE_ATLASCELLLABEL.Value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.Cell_CheckBox_Change"
End Sub

Private Sub Round_SpinButton_Change()
    On Error GoTo ErrorHandler
    If Not mbLocked And Round_SpinButton.Value <> CLng(ARESConfig.ARES_LENGTH_ROUND.Value) Then
        SetLocked True
        Round_Number_Label.Caption = Round_SpinButton.Value
        ARESConfig.ARES_LENGTH_ROUND.Value = CStr(Round_SpinButton.Value)
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.Round_SpinButton_Change"
End Sub

' ============================================================
' TRIGGER ID -> ARES_Length_Trigger_ID (also rewrites ARES_Length_Triggers)
' ============================================================

Private Sub Edit_Trigger_Command_Click()
    On Error GoTo ErrorHandler
    If Not mbLocked Then
        SetLocked True
        TextBox_Trigger.Value = ARESConfig.ARES_LENGTH_TRIGGER_ID.Value
        TextBox_Trigger.Visible = True
        Edit_Trigger_Command.Visible = False
        TextBox_Trigger.SetFocus
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.Edit_Trigger_Command_Click"
End Sub

Private Sub TextBox_Trigger_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    If TextBox_Trigger.Value <> ARESConfig.ARES_LENGTH_TRIGGER_ID.Value Then
        ARESConfig.ARES_LENGTH_TRIGGER.Value = Replace(ARESConfig.ARES_LENGTH_TRIGGER.Value, ARESConfig.ARES_LENGTH_TRIGGER_ID.Value, TextBox_Trigger.Value)
        ARESConfig.ARES_LENGTH_TRIGGER_ID.Value = TextBox_Trigger.Value
        Edit_Trigger_Command.Caption = GetTranslation("AutoLengthsGUIOptionsEdit_Trigger_CommandCaption", ARESConfig.ARES_LENGTH_TRIGGER_ID.Value)
    End If
    TextBox_Trigger.Visible = False
    Edit_Trigger_Command.Visible = True
    SetLocked False
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.TextBox_Trigger_Exit"
End Sub

Private Sub TextBox_Trigger_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    FormUXHelper.NoteInlineKeyDown KeyCode, Shift
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.TextBox_Trigger_KeyDown"
End Sub

Private Sub TextBox_Trigger_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            TextBox_Trigger_Exit returnB
            Edit_Trigger_Command.SetFocus
        Case FormUXKeyCancel
            FormUXHelper.RevertInlineEdit TextBox_Trigger, ARESConfig.ARES_LENGTH_TRIGGER_ID
            TextBox_Trigger_Exit returnB
            Edit_Trigger_Command.SetFocus
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.TextBox_Trigger_KeyUp"
End Sub

' ============================================================
' TRIGGERS LIST -> ARES_Length_Triggers (each entry must contain the trigger ID)
' ============================================================

Private Sub Edit_Triggers_List_Command_Click()
    On Error GoTo ErrorHandler
    If Not mbLocked Then
        SetLocked True
        TextBox_Triggers_List.Value = ARESConfig.ARES_LENGTH_TRIGGER.Value
        TextBox_Triggers_List.Visible = True
        Edit_Triggers_List_Command.Visible = False
        TextBox_Triggers_List.SetFocus
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.Edit_Triggers_List_Command_Click"
End Sub

Private Sub TextBox_Triggers_List_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    Dim txts() As String
    Dim i As Long

    If TextBox_Triggers_List.Value <> ARESConfig.ARES_LENGTH_TRIGGER.Value Then
        txts = Split(TextBox_Triggers_List.Value, ARESConstants.ARES_VAR_DELIMITER)
        For i = LBound(txts) To UBound(txts)
            If Not txts(i) Like "*" & ARESConfig.ARES_LENGTH_TRIGGER_ID.Value & "*" Then
                ShowStatus GetTranslation("AutoLengthsGUIOptionsEdit_Triggers_List_Error", ARESConfig.ARES_LENGTH_TRIGGER_ID.Value)
                Exit Sub          ' keep the editor open so the user can fix the invalid entry
            End If
        Next i
        ARESConfig.ARES_LENGTH_TRIGGER.Value = TextBox_Triggers_List.Value
    End If

    TextBox_Triggers_List.Visible = False
    Edit_Triggers_List_Command.Visible = True
    SetLocked False
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.TextBox_Triggers_List_Exit"
End Sub

Private Sub TextBox_Triggers_List_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    FormUXHelper.NoteInlineKeyDown KeyCode, Shift
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.TextBox_Triggers_List_KeyDown"
End Sub

Private Sub TextBox_Triggers_List_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            TextBox_Triggers_List_Exit returnB
            Edit_Triggers_List_Command.SetFocus
        Case FormUXKeyCancel
            FormUXHelper.RevertInlineEdit TextBox_Triggers_List, ARESConfig.ARES_LENGTH_TRIGGER
            TextBox_Triggers_List_Exit returnB
            Edit_Triggers_List_Command.SetFocus
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.TextBox_Triggers_List_KeyUp"
End Sub

' ============================================================
' ATLAS CELL LIST -> ARES_Cell_Is_Label_Name
' ============================================================

Private Sub Edit_Cells_List_Command_Click()
    On Error GoTo ErrorHandler
    If Not mbLocked Then
        SetLocked True
        TextBox_Cells_List.Value = ARESConfig.ARES_CELL_LIKE_LABEL.Value
        TextBox_Cells_List.Visible = True
        Edit_Cells_List_Command.Visible = False
        TextBox_Cells_List.SetFocus
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.Edit_Cells_List_Command_Click"
End Sub

Private Sub TextBox_Cells_List_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    FormUXHelper.CommitInlineEdit TextBox_Cells_List, Edit_Cells_List_Command, ARESConfig.ARES_CELL_LIKE_LABEL
    SetLocked False
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.TextBox_Cells_List_Exit"
End Sub

Private Sub TextBox_Cells_List_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    FormUXHelper.NoteInlineKeyDown KeyCode, Shift
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.TextBox_Cells_List_KeyDown"
End Sub

Private Sub TextBox_Cells_List_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            TextBox_Cells_List_Exit returnB
            Edit_Cells_List_Command.SetFocus
        Case FormUXKeyCancel
            FormUXHelper.RevertInlineEdit TextBox_Cells_List, ARESConfig.ARES_CELL_LIKE_LABEL
            TextBox_Cells_List_Exit returnB
            Edit_Cells_List_Command.SetFocus
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.TextBox_Cells_List_KeyUp"
End Sub

' ============================================================
' FORM LIFECYCLE
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption = GetTranslation("AutoLengthsGUIOptionsCaption")
    ' Checkbox captions live on the checkbox: Tab-focus visible + the text toggles the box
    Main_CheckBox.Caption = GetTranslation("AutoLengthsGUIOptionsMain_LabelCaption")
    Color_CheckBox.Caption = GetTranslation("AutoLengthsGUIOptionsColor_LabelCaption")
    Only_Color_CheckBox.Caption = GetTranslation("AutoLengthsGUIOptionsOnly_Color_LabelCaption")
    Cell_CheckBox.Caption = GetTranslation("AutoLengthsGUIOptionsCell_LabelCaption")
    Edit_Triggers_List_Command.Caption = GetTranslation("AutoLengthsGUIOptionsEdit_Triggers_List_CommandCaption")
    Round_Label.Caption = GetTranslation("AutoLengthsGUIOptionsRound_LabelCaption")
    Edit_Cells_List_Command.Caption = GetTranslation("AutoLengthsGUIOptionsEdit_Cells_List_CommandCaption")

    ' Tooltips
    FormUXHelper.SetTip Main_CheckBox, "AutoLengthsGUIOptionsMain_LabelTip"
    FormUXHelper.SetTip Color_CheckBox, "AutoLengthsGUIOptionsColor_LabelTip"
    FormUXHelper.SetTip Only_Color_CheckBox, "AutoLengthsGUIOptionsOnly_Color_LabelTip"
    FormUXHelper.SetTip Cell_CheckBox, "AutoLengthsGUIOptionsCell_LabelTip"
    FormUXHelper.SetTip Edit_Trigger_Command, "AutoLengthsGUIOptionsEdit_Trigger_CommandTip"
    FormUXHelper.SetTip Edit_Triggers_List_Command, "AutoLengthsGUIOptionsEdit_Triggers_List_CommandTip"
    FormUXHelper.SetTip Edit_Cells_List_Command, "AutoLengthsGUIOptionsEdit_Cells_List_CommandTip"
    FormUXHelper.SetTip Round_Label, "AutoLengthsGUIOptionsRound_LabelTip"
    FormUXHelper.SetTip Round_SpinButton, "AutoLengthsGUIOptionsRound_LabelTip"

    ' Restore-defaults button
    Reset_Command.Caption = GetTranslation("FormResetDefaultsCaption")
    FormUXHelper.SetTip Reset_Command, "FormResetDefaultsTip"

    SeedControls
    FormPlacement.RestoreFormPosition Me, Me.Name
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.UserForm_Initialize"
End Sub

' Re-seed all controls from the current config values.
Private Sub SeedControls()
    On Error GoTo ErrorHandler
    Edit_Trigger_Command.Caption = GetTranslation("AutoLengthsGUIOptionsEdit_Trigger_CommandCaption", ARESConfig.ARES_LENGTH_TRIGGER_ID.Value)

    Round_Number_Label.Caption = ARESConfig.ARES_LENGTH_ROUND.Value
    Round_SpinButton.Value = Round_Number_Label.Caption
    If ARESConfig.ARES_AUTO_LENGTHS.Value Then
        Main_CheckBox.Value = "True"
    Else
        Main_CheckBox.Value = "False"
    End If
    If ARESConfig.ARES_UPDATE_COLOR_WITH_LENGTH.Value Then
        Color_CheckBox.Value = "True"
    Else
        Color_CheckBox.Value = "False"
    End If
    If ARESConfig.ARES_ONLY_COLOR.Value Then
        Only_Color_CheckBox.Value = "True"
    Else
        Only_Color_CheckBox.Value = "False"
    End If
    If ARESConfig.ARES_UPDATE_ATLASCELLLABEL.Value Then
        Cell_CheckBox.Value = "True"
    Else
        Cell_CheckBox.Value = "False"
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.SeedControls"
End Sub

' Restore every option this form edits to its default value, persist, then re-seed.
Private Sub Reset_Command_Click()
    On Error GoTo ErrorHandler
    FormUXHelper.PersistDefault ARESConfig.ARES_AUTO_LENGTHS
    FormUXHelper.PersistDefault ARESConfig.ARES_UPDATE_COLOR_WITH_LENGTH
    FormUXHelper.PersistDefault ARESConfig.ARES_ONLY_COLOR
    FormUXHelper.PersistDefault ARESConfig.ARES_UPDATE_ATLASCELLLABEL
    FormUXHelper.PersistDefault ARESConfig.ARES_LENGTH_ROUND
    FormUXHelper.PersistDefault ARESConfig.ARES_LENGTH_TRIGGER_ID
    FormUXHelper.PersistDefault ARESConfig.ARES_LENGTH_TRIGGER
    FormUXHelper.PersistDefault ARESConfig.ARES_CELL_LIKE_LABEL
    SeedControls
    LangManager.ShowStatusT "FormDefaultsRestored"
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.Reset_Command_Click"
End Sub

' Any error path must call SetLocked False so controls are never left disabled.
Private Sub SetLocked(ByVal bState As Boolean)
    On Error GoTo ErrorHandler
    mbLocked = bState
    FormUXHelper.SetControlsLocked Me, bState
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.SetLocked"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHandler

    If mbLocked Then
        Cancel = True
        Select Case True
            Case TextBox_Triggers_List.Visible
                FormUXHelper.NudgeActiveEdit TextBox_Triggers_List
            Case TextBox_Cells_List.Visible
                FormUXHelper.NudgeActiveEdit TextBox_Cells_List
            Case TextBox_Trigger.Visible
                FormUXHelper.NudgeActiveEdit TextBox_Trigger
        End Select
    Else
        FormPlacement.SaveFormPosition Me, Me.Name
        command.OnAutoLengthsGUIClosed
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengths_GUI_Options.UserForm_QueryClose"
End Sub

