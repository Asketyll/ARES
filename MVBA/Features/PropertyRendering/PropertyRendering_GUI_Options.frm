VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PropertyRendering_GUI_Options 
   Caption         =   "PropertyRendering_GUI_Options"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3015
   OleObjectBlob   =   "PropertyRendering_GUI_Options.frx":0000
End
Attribute VB_Name = "PropertyRendering_GUI_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm: PropertyRendering_GUI_Options
' Description: Options panel for both text rendering (PropertyRendering: ARES_Text_Render + the two ATLAS
'              label-cell settings) and attribute actuation (PropertyActuator: ARES_Actuate_Color/Level).
'              SHARING IS THE PANEL ONLY: the two stay separate logic modules with separate doctrines and
'              pipeline call sites - do not let this shared panel become a reason to merge their logic.
'
'              OUTSTANDING MANUAL CLEANUP: the retired "Color_CheckBox" control (colour-sync option) and
'              the pilot-property picker controls of an earlier revision (ActuateColorProp_Label,
'              ComboBox_ActuateColorProp, ActuateLevelProp_Label, ComboBox_ActuateLevelProp), if present,
'              are unreferenced dead controls still in the designer/.frx - code cannot delete a visual
'              control, remove them manually in the VBA IDE.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass, ARESConfigClass, FormUXHelper, FormPlacement, Command,
'               PropertyActuator
Option Explicit

Private mbLocked As Boolean

' ============================================================
' MASTER SWITCH - CheckBox -> ARES_Text_Render
' ============================================================

Private Sub Main_CheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    ' Enter toggles the checkbox too (uniform with buttons; Space already toggles natively).
    If Shift = 0 And KeyCode = vbKeyReturn Then Main_CheckBox.value = Not Main_CheckBox.value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.Main_CheckBox_KeyUp"
End Sub

Private Sub Main_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(Main_CheckBox.value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_TEXT_RENDER.value <> sVal Then
        SetLocked True
        ARESConfig.ARES_TEXT_RENDER.value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.Main_CheckBox_Change"
End Sub

' ============================================================
' ATLAS LABEL REBUILD - CheckBox -> ARES_Update_ATLASCellLabel
' ============================================================

Private Sub Cell_CheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    If Shift = 0 And KeyCode = vbKeyReturn Then Cell_CheckBox.value = Not Cell_CheckBox.value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.Cell_CheckBox_KeyUp"
End Sub

Private Sub Cell_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(Cell_CheckBox.value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_UPDATE_ATLASCELLLABEL.value <> sVal Then
        SetLocked True
        ARESConfig.ARES_UPDATE_ATLASCELLLABEL.value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.Cell_CheckBox_Change"
End Sub

' ============================================================
' ATLAS CELL LIST -> ARES_Cell_Is_Label_Name
' Inline editor: the button hides and a TextBox takes its place, exactly as the Auto Lengths options form
' did - same FormUXHelper primitives, so the interaction is unchanged for the user.
' ============================================================

Private Sub Edit_Cells_List_Command_Click()
    On Error GoTo ErrorHandler
    If Not mbLocked Then
        SetLocked True
        TextBox_Cells_List.value = ARESConfig.ARES_CELL_LIKE_LABEL.value
        TextBox_Cells_List.Visible = True
        Edit_Cells_List_Command.Visible = False
        TextBox_Cells_List.SetFocus
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.Edit_Cells_List_Command_Click"
End Sub

Private Sub TextBox_Cells_List_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    FormUXHelper.CommitInlineEdit TextBox_Cells_List, Edit_Cells_List_Command, ARESConfig.ARES_CELL_LIKE_LABEL
    SetLocked False
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.TextBox_Cells_List_Exit"
End Sub

Private Sub TextBox_Cells_List_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    FormUXHelper.NoteInlineKeyDown KeyCode, Shift
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.TextBox_Cells_List_KeyDown"
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
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.TextBox_Cells_List_KeyUp"
End Sub

' ============================================================
' PROPERTY ACTUATOR - Color/Level attribute painting, hosted here (see module header). Two independent
' master switches only - pilot properties are fixed/reserved (ARES_Color/ARES_Lvl), nothing to pick.
' ============================================================

Private Sub ActuateColor_CheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    If Shift = 0 And KeyCode = vbKeyReturn Then ActuateColor_CheckBox.value = Not ActuateColor_CheckBox.value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.ActuateColor_CheckBox_KeyUp"
End Sub

Private Sub ActuateColor_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(ActuateColor_CheckBox.value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_ACTUATE_COLOR.value <> sVal Then
        SetLocked True
        ARESConfig.ARES_ACTUATE_COLOR.value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.ActuateColor_CheckBox_Change"
End Sub

Private Sub ActuateLevel_CheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    If Shift = 0 And KeyCode = vbKeyReturn Then ActuateLevel_CheckBox.value = Not ActuateLevel_CheckBox.value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.ActuateLevel_CheckBox_KeyUp"
End Sub

Private Sub ActuateLevel_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(ActuateLevel_CheckBox.value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_ACTUATE_LEVEL.value <> sVal Then
        SetLocked True
        ARESConfig.ARES_ACTUATE_LEVEL.value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.ActuateLevel_CheckBox_Change"
End Sub

' ============================================================
' FORM LIFECYCLE
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption = GetTranslation("RenderingGUIOptionsCaption")
    ' Checkbox captions live on the checkboxes: Tab-focus visible + the text toggles the box
    Main_CheckBox.Caption = GetTranslation("RenderingGUIOptionsMain_LabelCaption")
    Cell_CheckBox.Caption = GetTranslation("RenderingGUIOptionsCell_LabelCaption")
    Edit_Cells_List_Command.Caption = GetTranslation("RenderingGUIOptionsEdit_Cells_List_CommandCaption")
    ActuateColor_CheckBox.Caption = GetTranslation("RenderingGUIOptionsActuateColor_LabelCaption")
    ActuateLevel_CheckBox.Caption = GetTranslation("RenderingGUIOptionsActuateLevel_LabelCaption")
    ActuatorSection_Label.Caption = GetTranslation("RenderingGUIOptionsActuatorSection_LabelCaption")

    ' Tooltips
    FormUXHelper.SetTip Main_CheckBox, "RenderingGUIOptionsMain_LabelTip"
    FormUXHelper.SetTip Cell_CheckBox, "RenderingGUIOptionsCell_LabelTip"
    FormUXHelper.SetTip Edit_Cells_List_Command, "RenderingGUIOptionsEdit_Cells_List_CommandTip"
    FormUXHelper.SetTip ActuateColor_CheckBox, "RenderingGUIOptionsActuateColor_LabelTip"
    FormUXHelper.SetTip ActuateLevel_CheckBox, "RenderingGUIOptionsActuateLevel_LabelTip"

    ' Restore-defaults button
    Reset_Command.Caption = GetTranslation("FormResetDefaultsCaption")
    FormUXHelper.SetTip Reset_Command, "FormResetDefaultsTip"

    SeedControls
    FormPlacement.RestoreFormPosition Me, Me.Name
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.UserForm_Initialize"
End Sub

' Re-seed the checkboxes from the current config values. The cell-name list has no visible control
' until the user opens the inline editor, so nothing to seed for it here.
Private Sub SeedControls()
    On Error GoTo ErrorHandler

    Main_CheckBox.value = (UCase(Trim(ARESConfig.ARES_TEXT_RENDER.value)) = "TRUE")
    Cell_CheckBox.value = (UCase(Trim(ARESConfig.ARES_UPDATE_ATLASCELLLABEL.value)) = "TRUE")
    ActuateColor_CheckBox.value = (UCase(Trim(ARESConfig.ARES_ACTUATE_COLOR.value)) = "TRUE")
    ActuateLevel_CheckBox.value = (UCase(Trim(ARESConfig.ARES_ACTUATE_LEVEL.value)) = "TRUE")
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.SeedControls"
End Sub

' Restore every option this form edits to its default value, persist, then re-seed.
Private Sub Reset_Command_Click()
    On Error GoTo ErrorHandler
    If Not FormUXHelper.ConfirmReset() Then Exit Sub
    FormUXHelper.PersistDefault ARESConfig.ARES_TEXT_RENDER
    FormUXHelper.PersistDefault ARESConfig.ARES_UPDATE_ATLASCELLLABEL
    FormUXHelper.PersistDefault ARESConfig.ARES_CELL_LIKE_LABEL
    FormUXHelper.PersistDefault ARESConfig.ARES_ACTUATE_COLOR
    FormUXHelper.PersistDefault ARESConfig.ARES_ACTUATE_LEVEL
    PropertyActuator.RefreshActuatorState
    SeedControls
    LangManager.ShowStatusT "FormDefaultsRestored"
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.Reset_Command_Click"
End Sub

' Any error path must call SetLocked False so controls are never left disabled.
Private Sub SetLocked(ByVal bState As Boolean)
    On Error GoTo ErrorHandler
    mbLocked = bState
    FormUXHelper.SetControlsLocked Me, bState
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.SetLocked"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHandler

    If mbLocked Then
        Cancel = True
    Else
        FormPlacement.SaveFormPosition Me, Me.Name
        command.OnPropertyRenderingGUIClosed
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.UserForm_QueryClose"
End Sub



