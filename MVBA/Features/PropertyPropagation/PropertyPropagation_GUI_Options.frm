VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PropertyPropagation_GUI_Options 
   Caption         =   "PropertyPropagation_GUI_Options"
   ClientHeight    =   1455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4815
   OleObjectBlob   =   "PropertyPropagation_GUI_Options.frx":0000
End
Attribute VB_Name = "PropertyPropagation_GUI_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm: PropertyPropagation_GUI_Options
' Description: Options panel for Property Propagation - the value-calc master switch
'              (ARES_Property_Propagation) and the detach-empty option (ARES_Propagation_Detach_Empty).
'              The trigger cells and target properties now come from the @cell rules in Property Tagging
'              (GUI 1, epic 12) - this form owns no cell-name / target-property config.
'
'              DESIGNER (manual, Asketyll) - controls required with EXACTLY these names:
'                Main_CheckBox (CheckBox, value master), DetachEmpty_CheckBox (CheckBox, detach-empty
'                option; caption set in code), Reset_Command (CommandButton).
'              StartUpPosition = 0 Manual. Tab order: master -> detach-empty -> reset.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass, ARESConfigClass, FormUXHelper, FormPlacement, Command
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
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.Main_CheckBox_Change"
End Sub

' ============================================================
' DETACH-EMPTY OPTION - CheckBox -> ARES_Propagation_Detach_Empty (round-4)
' When on, an emptied value is DETACHED (via the tagger) instead of cleared. Independent of the master
' switch (it may be on while the master is off).
' ============================================================

Private Sub DetachEmpty_CheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    ' Enter toggles the checkbox too (uniform with buttons; Space already toggles natively).
    If Shift = 0 And KeyCode = vbKeyReturn Then DetachEmpty_CheckBox.value = Not DetachEmpty_CheckBox.value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.DetachEmpty_CheckBox_KeyUp"
End Sub

Private Sub DetachEmpty_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(DetachEmpty_CheckBox.value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_PROPAGATION_DETACH_EMPTY.value <> sVal Then
        SetLocked True
        ARESConfig.ARES_PROPAGATION_DETACH_EMPTY.value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.DetachEmpty_CheckBox_Change"
End Sub

' ============================================================
' FORM LIFECYCLE
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption = GetTranslation("PropagationGUIOptionsCaption")
    ' Checkbox captions live on the checkboxes: Tab-focus visible + the text toggles the box
    Main_CheckBox.Caption = GetTranslation("PropagationGUIOptionsMain_LabelCaption")
    DetachEmpty_CheckBox.Caption = GetTranslation("PropagationGUIOptionsDetachEmpty_LabelCaption")

    ' Tooltips
    FormUXHelper.SetTip Main_CheckBox, "PropagationGUIOptionsMain_LabelTip"
    FormUXHelper.SetTip DetachEmpty_CheckBox, "PropagationGUIOptionsDetachEmpty_LabelTip"

    ' Restore-defaults button
    Reset_Command.Caption = GetTranslation("FormResetDefaultsCaption")
    FormUXHelper.SetTip Reset_Command, "FormResetDefaultsTip"

    SeedControls
    FormPlacement.RestoreFormPosition Me, Me.Name
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.UserForm_Initialize"
End Sub

' Re-seed the two checkboxes from the current config values.
Private Sub SeedControls()
    On Error GoTo ErrorHandler

    Main_CheckBox.value = (UCase(Trim(ARESConfig.ARES_PROPERTY_PROPAGATION.value)) = "TRUE")
    ' Detach-empty option is independent of the master switch (seeded like Main_CheckBox).
    DetachEmpty_CheckBox.value = (UCase(Trim(ARESConfig.ARES_PROPAGATION_DETACH_EMPTY.value)) = "TRUE")
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.SeedControls"
End Sub

' Restore every option this form edits to its default value, persist, then re-seed.
Private Sub Reset_Command_Click()
    On Error GoTo ErrorHandler
    FormUXHelper.PersistDefault ARESConfig.ARES_PROPERTY_PROPAGATION
    FormUXHelper.PersistDefault ARESConfig.ARES_PROPAGATION_DETACH_EMPTY
    SeedControls
    LangManager.ShowStatusT "FormDefaultsRestored"
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.Reset_Command_Click"
End Sub

' Any error path must call SetLocked False so controls are never left disabled.
Private Sub SetLocked(ByVal bState As Boolean)
    On Error GoTo ErrorHandler
    mbLocked = bState
    FormUXHelper.SetControlsLocked Me, bState
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.SetLocked"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHandler

    If mbLocked Then
        Cancel = True
    Else
        FormPlacement.SaveFormPosition Me, Me.Name
        command.OnPropagationGUIClosed
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyPropagation_GUI_Options.UserForm_QueryClose"
End Sub


