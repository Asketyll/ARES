VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PropertyRendering_GUI_Options 
   Caption         =   "PropertyRendering_GUI_Options"
   ClientHeight    =   4455
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
' Description: Options panel for EVERYTHING THAT RENDERS A CUSTOM PROPERTY VISIBLE - text or graphic
'              attribute (Asketyll, 2026-08-19: "dans le formulaire de rendu, ce n'est pas du texte mais
'              c'est du rendu" - painting a Color/Level FROM a property is conceptually the same family of
'              visible result as a Prop[Name] text token, even though it is a SEPARATE engine/doctrine).
'              This panel hosts the controls of TWO independent Depth-0 engines:
'                1. PropertyRendering (text) - the render master switch (ARES_Text_Render) and the three
'                   text-presentation options inherited from the removed Auto Lengths form: colour sync
'                   (ARES_Only_Color_Update) and the two ATLAS label-cell settings
'                   (ARES_Update_ATLASCellLabel, ARES_Cell_Is_Label_Name).
'                2. PropertyActuator (attribute, epic 16) - two independent master switches only
'                   (ARES_Actuate_Color, ARES_Actuate_Level). Pilot properties are FIXED and RESERVED
'                   (ARES_Color/ARES_Lvl, revised 2026-08-20 after a real-world test exposed a silent-error
'                   class in an earlier configurable-picker design - see PropertyActuator's own header) -
'                   no picker, nothing to configure beyond the two switches.
'              SHARING IS THE PANEL ONLY, NOT THE MODULES: PropertyRendering.bas and PropertyActuator.bas
'              stay two separate logic modules with two separate doctrines (Rendering writes TEXT only,
'              Actuator writes ATTRIBUTES only) and two separate Depth-0 pipeline call sites in
'              ElementChangeHandler.cls (Actuator runs BEFORE Rendering - see PropertyActuator's own header).
'              Do NOT let this shared panel become a reason to merge the two modules' logic - a future
'              agent editing THIS FILE is touching UI wiring for two engines at once; a future agent editing
'              PropertyActuator.bas or PropertyRendering.bas is touching exactly one engine, as before.
'
'              WHY HERE AND NOT IN THE CALCULATION FORM (the three original text-presentation options): all
'              three serve DISPLAY, not value computation. CellRedreaw - the sole consumer of the two ATLAS
'              settings - is called by StringsInEl.SetTextAtSubId, which is the renderer's only text write.
'              Housing them under Calculation would have put them one engine away from the code that reads
'              them. ARES_Length_Round stays with PropertyCalculation_GUI_Options: it is the default
'              decimals of the Length/GroupLength calc SOURCES, which is literally its semantics.
'
'              The colour-sync option (ARES_Only_Color_Update) is hosted here as the least-wrong home while
'              the legacy colour hook lives on in ElementChangeHandler. When the property-driven colour
'              mechanism (PropertyActuator) replaces that hook, this checkbox moves or dies with it - it is
'              NOT a statement that colour sync belongs to the renderer, and this story does NOT retire the
'              legacy hook (separate, later story per the actuator's own cahier des charges �5).
'
'              DESIGNER (manual) - controls required with EXACTLY these names:
'                Main_CheckBox (CheckBox, render master switch), Color_CheckBox (CheckBox, colour sync),
'                Cell_CheckBox (CheckBox, ATLAS label rebuild), TextBox_Cells_List (TextBox, Visible = False
'                in the designer - the inline editor for the cell-name list), Edit_Cells_List_Command
'                (CommandButton, Visible = True - swaps with the TextBox),
'                ActuatorSection_Label (Label, bold/section-header style - "Property Actuator" divider so
'                the two engines' controls are never visually mistaken for one group; reviewer-4 flagged
'                this labelling as the one thing to get right about sharing this panel),
'                ActuateColor_CheckBox (CheckBox, PropertyActuator Color master switch),
'                ActuateLevel_CheckBox (CheckBox, PropertyActuator Level master switch),
'                Reset_Command (CommandButton).
'              REVISED 2026-08-20: the pilot-property picker controls (ActuateColorProp_Label,
'              ComboBox_ActuateColorProp, ActuateLevelProp_Label, ComboBox_ActuateLevelProp) that an earlier
'              revision of this panel required are DROPPED - pilot properties are now fixed/reserved
'              (ARES_Color/ARES_Lvl), nothing to pick. If any of those 4 controls were already added in the
'              designer, they can be deleted - the code no longer references them.
'              VISUAL GROUPING (manual, designer): place ActuatorSection_Label + its 2 checkboxes in their
'              own block, separated from the render/colour-sync/ATLAS controls above by visible whitespace or
'              a horizontal rule, so the panel reads as two clearly bounded sections under one window, not one
'              flat list where an Actuator checkbox could be mistaken for a Rendering option or vice versa.
'              NO help button here, unlike the Tagging and Calculation panels: those exist to open the wiki
'              because their rule grammars do not fit in a tooltip. The actuator has no grammar to explain
'              beyond its two switches - every control is explained by its own tooltip.
'              StartUpPosition = 0 Manual. Tab order: master -> colour -> cell -> edit-cells ->
'              actuate-color -> actuate-level -> reset.
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
' PROPERTY ACTUATOR (epic 16) - Color/Level attribute painting, hosted here per Asketyll's ruling (see
' module header). Two independent master switches only - pilot properties are fixed/reserved
' (ARES_Color/ARES_Lvl), revised 2026-08-20, nothing to pick.
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
    Color_CheckBox.Caption = GetTranslation("RenderingGUIOptionsColor_LabelCaption")
    Cell_CheckBox.Caption = GetTranslation("RenderingGUIOptionsCell_LabelCaption")
    Edit_Cells_List_Command.Caption = GetTranslation("RenderingGUIOptionsEdit_Cells_List_CommandCaption")
    ActuateColor_CheckBox.Caption = GetTranslation("RenderingGUIOptionsActuateColor_LabelCaption")
    ActuateLevel_CheckBox.Caption = GetTranslation("RenderingGUIOptionsActuateLevel_LabelCaption")
    ActuatorSection_Label.Caption = GetTranslation("RenderingGUIOptionsActuatorSection_LabelCaption")

    ' Tooltips
    FormUXHelper.SetTip Main_CheckBox, "RenderingGUIOptionsMain_LabelTip"
    FormUXHelper.SetTip Color_CheckBox, "RenderingGUIOptionsColor_LabelTip"
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

' Re-seed the three checkboxes from the current config values. The cell-name list has no visible control
' until the user opens the inline editor, so nothing to seed for it here.
Private Sub SeedControls()
    On Error GoTo ErrorHandler

    Main_CheckBox.value = (UCase(Trim(ARESConfig.ARES_TEXT_RENDER.value)) = "TRUE")
    Color_CheckBox.value = (UCase(Trim(ARESConfig.ARES_ONLY_COLOR.value)) = "TRUE")
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
    FormUXHelper.PersistDefault ARESConfig.ARES_ONLY_COLOR
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


