VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PropertyRendering_GUI_Options 
   Caption         =   "PropertyRendering_GUI_Options"
   ClientHeight    =   2175
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
' Description: Options panel for Property Rendering - the third engine's own options panel, which it did
'              not have. It carries the render master switch (ARES_Text_Render, previously editable from NO
'              form at all) and the three text-presentation options that Auto Lengths' options form used to
'              host and that OUTLIVE it: colour sync (ARES_Only_Color_Update) and the two ATLAS label-cell
'              settings (ARES_Update_ATLASCellLabel, ARES_Cell_Is_Label_Name).
'
'              WHY HERE AND NOT IN THE CALCULATION FORM: all three serve DISPLAY, not value computation.
'              CellRedreaw - the sole consumer of the two ATLAS settings - is called by
'              StringsInEl.SetTextAtSubId, which is the renderer's only text write. Housing them under
'              Calculation would have put them one engine away from the code that reads them.
'              ARES_Length_Round stays with PropertyCalculation_GUI_Options: it is the default decimals of
'              the Length/GroupLength calc SOURCES, which is literally its semantics.
'
'              The colour-sync option is hosted here as the least-wrong home while the legacy colour hook
'              lives on in ElementChangeHandler. When the property-driven colour mechanism replaces that
'              hook, this checkbox moves or dies with it - it is NOT a statement that colour sync belongs
'              to the renderer.
'
'              DESIGNER (manual) - controls required with EXACTLY these names:
'                Main_CheckBox (CheckBox, render master switch), Color_CheckBox (CheckBox, colour sync),
'                Cell_CheckBox (CheckBox, ATLAS label rebuild), TextBox_Cells_List (TextBox, Visible = False
'                in the designer - the inline editor for the cell-name list), Edit_Cells_List_Command
'                (CommandButton, Visible = True - swaps with the TextBox), Reset_Command (CommandButton).
'              NO help button here, unlike the Tagging and Calculation panels: those exist to open the wiki
'              because their rule grammars do not fit in a tooltip. This panel has no rules - three
'              checkboxes and a "|"-separated name list - so every control is explained by its own tooltip.
'              StartUpPosition = 0 Manual. Tab order: master -> colour -> cell -> edit-cells -> reset.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass, ARESConfigClass, FormUXHelper, FormPlacement, Command
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
' COLOUR SYNC - CheckBox -> ARES_Only_Color_Update
' Independent of the master switch: the colour hook lives in ElementChangeHandler and runs whether or not
' the renderer is on.
' ============================================================

Private Sub Color_CheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    If Shift = 0 And KeyCode = vbKeyReturn Then Color_CheckBox.value = Not Color_CheckBox.value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.Color_CheckBox_KeyUp"
End Sub

Private Sub Color_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(Color_CheckBox.value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_ONLY_COLOR.value <> sVal Then
        SetLocked True
        ARESConfig.ARES_ONLY_COLOR.value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.Color_CheckBox_Change"
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
' FORM LIFECYCLE
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption = GetTranslation("RenderingGUIOptionsCaption")
    ' Checkbox captions live on the checkboxes: Tab-focus visible + the text toggles the box
    Main_CheckBox.Caption = GetTranslation("RenderingGUIOptionsMain_LabelCaption")
    Color_CheckBox.Caption = GetTranslation("RenderingGUIOptionsColor_LabelCaption")
    Cell_CheckBox.Caption = GetTranslation("RenderingGUIOptionsCell_LabelCaption")
    Edit_Cells_List_Command.Caption = GetTranslation("RenderingGUIOptionsEdit_Cells_List_CommandCaption")

    ' Tooltips
    FormUXHelper.SetTip Main_CheckBox, "RenderingGUIOptionsMain_LabelTip"
    FormUXHelper.SetTip Color_CheckBox, "RenderingGUIOptionsColor_LabelTip"
    FormUXHelper.SetTip Cell_CheckBox, "RenderingGUIOptionsCell_LabelTip"
    FormUXHelper.SetTip Edit_Cells_List_Command, "RenderingGUIOptionsEdit_Cells_List_CommandTip"

    ' Restore-defaults button
    Reset_Command.Caption = GetTranslation("FormResetDefaultsCaption")
    FormUXHelper.SetTip Reset_Command, "FormResetDefaultsTip"

    SeedControls
    FormPlacement.RestoreFormPosition Me, Me.Name
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.UserForm_Initialize"
End Sub

' Re-seed the three checkboxes from the current config values. The cell-name list has no visible control
' until the user opens the inline editor, so nothing to seed for it here.
Private Sub SeedControls()
    On Error GoTo ErrorHandler

    Main_CheckBox.value = (UCase(Trim(ARESConfig.ARES_TEXT_RENDER.value)) = "TRUE")
    Color_CheckBox.value = (UCase(Trim(ARESConfig.ARES_ONLY_COLOR.value)) = "TRUE")
    Cell_CheckBox.value = (UCase(Trim(ARESConfig.ARES_UPDATE_ATLASCELLLABEL.value)) = "TRUE")
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRendering_GUI_Options.SeedControls"
End Sub

' Restore every option this form edits to its default value, persist, then re-seed.
Private Sub Reset_Command_Click()
    On Error GoTo ErrorHandler
    If Not FormUXHelper.ConfirmReset() Then Exit Sub
    FormUXHelper.PersistDefault ARESConfig.ARES_TEXT_RENDER
    FormUXHelper.PersistDefault ARESConfig.ARES_ONLY_COLOR
    FormUXHelper.PersistDefault ARESConfig.ARES_UPDATE_ATLASCELLLABEL
    FormUXHelper.PersistDefault ARESConfig.ARES_CELL_LIKE_LABEL
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

