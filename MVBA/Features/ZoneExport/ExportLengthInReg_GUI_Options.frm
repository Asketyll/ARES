VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExportLengthInReg_GUI_Options 
   Caption         =   "Edit export length in region options:"
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3015
   OleObjectBlob   =   "ExportLengthInReg_GUI_Options.frx":0000
   StartUpPosition =   0  'Manual
End
Attribute VB_Name = "ExportLengthInReg_GUI_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm: ExportLengthInReg_GUI_Options
' Description: Options panel for ExportLengthInRegion - zone level, grouping key, rounding, save dialog.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ARESConfigClass, ErrorHandlerClass, FormUXHelper
Option Explicit

Private mbLocked As Boolean

' ============================================================
' ZONE LEVEL - Edit button + hidden TextBox
' ============================================================

Private Sub Edit_Level_Region_Command_Click()
    On Error GoTo ErrorHandler
    If Not mbLocked Then
        SetLocked True
        TextBox_RegionLevel.Value = ARESConfig.ARES_ZONING_OUTPUT_LEVEL.Value
        TextBox_RegionLevel.Visible = True
        Edit_Level_Region_Command.Visible = False
        TextBox_RegionLevel.SetFocus
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.Edit_Level_Region_Command_Click"
End Sub

Private Sub TextBox_RegionLevel_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = Trim(TextBox_RegionLevel.Value)
    If Len(sVal) > 0 And sVal <> ARESConfig.ARES_ZONING_OUTPUT_LEVEL.Value Then
        ARESConfig.ARES_ZONING_OUTPUT_LEVEL.Value = sVal
    End If
    TextBox_RegionLevel.Visible = False
    Edit_Level_Region_Command.Caption = GetTranslation("ZoneExportGUIOptionsEdit_Level_Region_CommandCaption")
    Edit_Level_Region_Command.Visible = True
    SetLocked False
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.TextBox_RegionLevel_Exit"
End Sub

Private Sub TextBox_RegionLevel_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    FormUXHelper.NoteInlineKeyDown KeyCode, Shift
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.TextBox_RegionLevel_KeyDown"
End Sub

Private Sub TextBox_RegionLevel_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            TextBox_RegionLevel_Exit returnB
            Edit_Level_Region_Command.SetFocus
        Case FormUXKeyCancel
            FormUXHelper.RevertInlineEdit TextBox_RegionLevel, ARESConfig.ARES_ZONING_OUTPUT_LEVEL
            TextBox_RegionLevel_Exit returnB
            Edit_Level_Region_Command.SetFocus
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.TextBox_RegionLevel_KeyUp"
End Sub

' ============================================================
' GROUP BY - ComboBox (localized display, stable English stored key)
' ============================================================

Private Sub ComboBox_Export_Type_Change()
    On Error GoTo ErrorHandler
    If mbLocked Then Exit Sub
    Dim sKey As String
    sKey = GroupByKeyFromDisplay()
    If ARESConfig.ARES_ZONE_EXPORT_GROUP_BY.Value <> sKey Then
        SetLocked True
        ARESConfig.ARES_ZONE_EXPORT_GROUP_BY.Value = sKey
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.ComboBox_Export_Type_Change"
End Sub

' ============================================================
' ROUNDING - SpinButton
' ============================================================

Private Sub Round_SpinButton_Change()
    On Error GoTo ErrorHandler
    If Not mbLocked And CStr(Round_SpinButton.Value) <> ARESConfig.ARES_ZONE_EXPORT_ROUND.Value Then
        SetLocked True
        Round_Number_Label.Caption = Round_SpinButton.Value
        ARESConfig.ARES_ZONE_EXPORT_ROUND.Value = CStr(Round_SpinButton.Value)
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.Round_SpinButton_Change"
End Sub

' ============================================================
' USE DIALOG - CheckBox
' ============================================================

Private Sub Use_Dialog_CheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    ' Enter toggles the checkbox too (uniform with buttons; Space already toggles natively).
    If Shift = 0 And KeyCode = vbKeyReturn Then Use_Dialog_CheckBox.Value = Not Use_Dialog_CheckBox.Value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.Use_Dialog_CheckBox_KeyUp"
End Sub

Private Sub Use_Dialog_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(Use_Dialog_CheckBox.Value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_ZONE_EXPORT_USE_DIALOG.Value <> sVal Then
        SetLocked True
        ARESConfig.ARES_ZONE_EXPORT_USE_DIALOG.Value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.Use_Dialog_CheckBox_Change"
End Sub

' ============================================================
' INITIALIZATION
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption = GetTranslation("ZoneExportGUIOptionsCaption")
    Round_Label.Caption = GetTranslation("ZoneExportGUIOptionsRound_LabelCaption")
    ' Checkbox caption lives on the checkbox: Tab-focus visible + the text toggles the box
    Use_Dialog_CheckBox.Caption = GetTranslation("ZoneExportGUIOptionsUse_Dialog_LabelCaption")
    Edit_Level_Region_Command.Caption = GetTranslation("ZoneExportGUIOptionsEdit_Level_Region_CommandCaption")
    GroupBy_Label.Caption = GetTranslation("ZoneExportGUIOptionsGroupBy_LabelCaption")

    ' Tooltips (AC-6)
    FormUXHelper.SetTip Edit_Level_Region_Command, "ZoneExportGUIOptionsEdit_Level_Region_CommandTip"
    FormUXHelper.SetTip GroupBy_Label, "ZoneExportGUIOptionsGroupBy_LabelTip"
    FormUXHelper.SetTip ComboBox_Export_Type, "ZoneExportGUIOptionsGroupBy_LabelTip"
    FormUXHelper.SetTip Round_Label, "ZoneExportGUIOptionsRound_LabelTip"
    FormUXHelper.SetTip Round_SpinButton, "ZoneExportGUIOptionsRound_LabelTip"
    FormUXHelper.SetTip Use_Dialog_CheckBox, "ZoneExportGUIOptionsUse_Dialog_LabelTip"

    ' Restore-defaults button
    Reset_Command.Caption = GetTranslation("FormResetDefaultsCaption")
    FormUXHelper.SetTip Reset_Command, "FormResetDefaultsTip"

    ' Group-by combo items: localized display, stable English stored key
    ComboBox_Export_Type.Clear
    ComboBox_Export_Type.AddItem GroupByDisplayFromKey("Style")
    ComboBox_Export_Type.AddItem GroupByDisplayFromKey("Level")
    ComboBox_Export_Type.AddItem GroupByDisplayFromKey("Color")

    ' Rounding spin bounds (value seeded in SeedControls, guarded against non-numeric config)
    Round_SpinButton.Min = 0
    Round_SpinButton.Max = 10

    SeedControls
    FormPlacement.RestoreFormPosition Me, Me.Name
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.UserForm_Initialize"
End Sub

' ============================================================
' CLOSE
' ============================================================

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHandler
    If mbLocked Then
        Cancel = True
        If TextBox_RegionLevel.Visible Then FormUXHelper.NudgeActiveEdit TextBox_RegionLevel
    Else
        FormPlacement.SaveFormPosition Me, Me.Name
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.UserForm_QueryClose"
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    command.OnZoneExportGUIClosed
End Sub

' ============================================================
' HELPERS
' ============================================================

' Re-seed all controls from the current config values.
Private Sub SeedControls()
    On Error GoTo ErrorHandler
    Dim sGroupBy As String
    sGroupBy = Trim(ARESConfig.ARES_ZONE_EXPORT_GROUP_BY.Value)
    If sGroupBy <> "Level" And sGroupBy <> "Color" Then sGroupBy = "Style"
    ComboBox_Export_Type.Value = GroupByDisplayFromKey(sGroupBy)

    Dim nRound As Integer
    If IsNumeric(ARESConfig.ARES_ZONE_EXPORT_ROUND.Value) Then
        nRound = CInt(ARESConfig.ARES_ZONE_EXPORT_ROUND.Value)
    Else
        nRound = CInt(ARESConfig.ARES_ZONE_EXPORT_ROUND.DefaultValue)
    End If
    If nRound < 0 Then nRound = 0
    If nRound > 10 Then nRound = 10
    Round_SpinButton.Value = nRound
    Round_Number_Label.Caption = nRound

    Use_Dialog_CheckBox.Value = (UCase(Trim(ARESConfig.ARES_ZONE_EXPORT_USE_DIALOG.Value)) = "TRUE")

    TextBox_RegionLevel.Visible = False
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.SeedControls"
End Sub

' NOTE: ARES_Zoning_Output_Level is the region output level shown here and is shared with the Zoning form.
Private Sub Reset_Command_Click()
    On Error GoTo ErrorHandler
    FormUXHelper.PersistDefault ARESConfig.ARES_ZONING_OUTPUT_LEVEL
    FormUXHelper.PersistDefault ARESConfig.ARES_ZONE_EXPORT_GROUP_BY
    FormUXHelper.PersistDefault ARESConfig.ARES_ZONE_EXPORT_ROUND
    FormUXHelper.PersistDefault ARESConfig.ARES_ZONE_EXPORT_USE_DIALOG
    SeedControls
    LangManager.ShowStatusT "FormDefaultsRestored"
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.Reset_Command_Click"
End Sub

' Explicit-state lock: replaces the toggle Locked()/CheckControlForLock pair.
Private Sub SetLocked(ByVal bState As Boolean)
    On Error GoTo ErrorHandler
    mbLocked = bState
    FormUXHelper.SetControlsLocked Me, bState
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.SetLocked"
End Sub

' Map the localized combo display back to the stable English stored key.
Private Function GroupByKeyFromDisplay() As String
    On Error GoTo ErrorHandler
    Dim sDisp As String
    sDisp = ComboBox_Export_Type.Value
    If sDisp = GetTranslation("ZoneExportGroupByLevel") Then
        GroupByKeyFromDisplay = "Level"
    ElseIf sDisp = GetTranslation("ZoneExportGroupByColor") Then
        GroupByKeyFromDisplay = "Color"
    Else
        GroupByKeyFromDisplay = "Style"
    End If
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.GroupByKeyFromDisplay"
    GroupByKeyFromDisplay = "Style"
End Function

' Map a stable English stored key to its localized combo display.
Private Function GroupByDisplayFromKey(ByVal sKey As String) As String
    On Error GoTo ErrorHandler
    Select Case sKey
        Case "Level"
            GroupByDisplayFromKey = GetTranslation("ZoneExportGroupByLevel")
        Case "Color"
            GroupByDisplayFromKey = GetTranslation("ZoneExportGroupByColor")
        Case Else
            GroupByDisplayFromKey = GetTranslation("ZoneExportGroupByStyle")
    End Select
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.GroupByDisplayFromKey"
    GroupByDisplayFromKey = GetTranslation("ZoneExportGroupByStyle")
End Function
