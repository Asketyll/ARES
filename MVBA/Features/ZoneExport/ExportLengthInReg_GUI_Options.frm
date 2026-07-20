VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExportLengthInReg_GUI_Options 
   Caption         =   "Edit export length in region options:"
   ClientHeight    =   4095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3615
   OleObjectBlob   =   "ExportLengthInReg_GUI_Options.frx":0000
End
Attribute VB_Name = "ExportLengthInReg_GUI_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm: ExportLengthInReg_GUI_Options
' Description: Options panel for ExportLengthInRegion - zone level, candidate level filter, grouping key, rounding, save dialog.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ARESConfigClass, ErrorHandlerClass, FormUXHelper, CustomPropertyHandler
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
' CANDIDATE LEVEL FILTER - Edit button + hidden TextBox
' Restricts the measured elements to these level(s); empty = all levels.
' Distinct from the ZONE level above (ARES_ZONING_OUTPUT_LEVEL).
' ============================================================

Private Sub Edit_Level_Candidate_Command_Click()
    On Error GoTo ErrorHandler
    If Not mbLocked Then
        SetLocked True
        TextBox_CandidateLevel.Value = ARESConfig.ARES_ZONE_EXPORT_LEVEL.Value
        TextBox_CandidateLevel.Visible = True
        Edit_Level_Candidate_Command.Visible = False
        TextBox_CandidateLevel.SetFocus
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.Edit_Level_Candidate_Command_Click"
End Sub

Private Sub TextBox_CandidateLevel_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    ' Empty is a valid value here (clears the filter -> all levels), so no Len>0 guard.
    Dim sVal As String
    sVal = Trim(TextBox_CandidateLevel.Value)
    If sVal <> ARESConfig.ARES_ZONE_EXPORT_LEVEL.Value Then
        ARESConfig.ARES_ZONE_EXPORT_LEVEL.Value = sVal
    End If
    TextBox_CandidateLevel.Visible = False
    Edit_Level_Candidate_Command.Caption = GetTranslation("ZoneExportGUIOptionsEdit_Level_Candidate_CommandCaption")
    Edit_Level_Candidate_Command.Visible = True
    SetLocked False
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.TextBox_CandidateLevel_Exit"
End Sub

Private Sub TextBox_CandidateLevel_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    FormUXHelper.NoteInlineKeyDown KeyCode, Shift
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.TextBox_CandidateLevel_KeyDown"
End Sub

Private Sub TextBox_CandidateLevel_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            TextBox_CandidateLevel_Exit returnB
            Edit_Level_Candidate_Command.SetFocus
        Case FormUXKeyCancel
            FormUXHelper.RevertInlineEdit TextBox_CandidateLevel, ARESConfig.ARES_ZONE_EXPORT_LEVEL
            TextBox_CandidateLevel_Exit returnB
            Edit_Level_Candidate_Command.SetFocus
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.TextBox_CandidateLevel_KeyUp"
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
' PER-ZONE SPLIT - CheckBox (turns on the per-zone breakdown)
' ============================================================

Private Sub CheckBox_PerZone_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    ' Enter toggles the checkbox too (uniform with buttons; Space already toggles natively).
    If Shift = 0 And KeyCode = vbKeyReturn Then CheckBox_PerZone.Value = Not CheckBox_PerZone.Value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.CheckBox_PerZone_KeyUp"
End Sub

Private Sub CheckBox_PerZone_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(CheckBox_PerZone.Value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_ZONE_EXPORT_PER_ZONE.Value <> sVal Then
        SetLocked True
        ARESConfig.ARES_ZONE_EXPORT_PER_ZONE.Value = sVal
        SetLocked False
    End If
    ApplyPerZoneEnabled
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.CheckBox_PerZone_Change"
End Sub

' ============================================================
' ZONE PROPERTY - ComboBox (custom-property read ON EACH ZONE for its label)
' Populated from ARES_Custom_Property_List so only valid names are selectable.
' ============================================================

Private Sub ComboBox_ZoneProperty_Change()
    On Error GoTo ErrorHandler
    If mbLocked Then Exit Sub
    Dim sVal As String
    ' Null-safe read: a dropdown-list combo with no selection returns Null (assigning it to a
    ' String would raise Error 94). Nested If (not And) per the no-short-circuit cheatsheet rule.
    If IsNull(ComboBox_ZoneProperty.Value) Then
        sVal = ""
    Else
        sVal = Trim(CStr(ComboBox_ZoneProperty.Value))
    End If
    If ARESConfig.ARES_ZONE_EXPORT_ZONE_PROPERTY.Value <> sVal Then
        SetLocked True
        ARESConfig.ARES_ZONE_EXPORT_ZONE_PROPERTY.Value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.ComboBox_ZoneProperty_Change"
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
' OPEN AFTER EXPORT - CheckBox (surfaces ARES_Zone_Export_Excel_Visible)
' When on, the saved workbook is left visible in Excel after the export.
' ============================================================

Private Sub OpenAfter_CheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    ' Enter toggles the checkbox too (uniform with buttons; Space already toggles natively).
    If Shift = 0 And KeyCode = vbKeyReturn Then OpenAfter_CheckBox.Value = Not OpenAfter_CheckBox.Value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.OpenAfter_CheckBox_KeyUp"
End Sub

Private Sub OpenAfter_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(OpenAfter_CheckBox.Value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_ZONE_EXPORT_EXCEL_VISIBLE.Value <> sVal Then
        SetLocked True
        ARESConfig.ARES_ZONE_EXPORT_EXCEL_VISIBLE.Value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.OpenAfter_CheckBox_Change"
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
    OpenAfter_CheckBox.Caption = GetTranslation("ZoneExportGUIOptionsOpenAfter_LabelCaption")
    CheckBox_PerZone.Caption = GetTranslation("ZoneExportGUIOptionsPerZone_LabelCaption")
    Edit_Level_Region_Command.Caption = GetTranslation("ZoneExportGUIOptionsEdit_Level_Region_CommandCaption")
    Edit_Level_Candidate_Command.Caption = GetTranslation("ZoneExportGUIOptionsEdit_Level_Candidate_CommandCaption")
    GroupBy_Label.Caption = GetTranslation("ZoneExportGUIOptionsGroupBy_LabelCaption")
    ZoneProperty_Label.Caption = GetTranslation("ZoneExportGUIOptionsZoneProperty_LabelCaption")

    ' Tooltips (AC-6)
    FormUXHelper.SetTip Edit_Level_Region_Command, "ZoneExportGUIOptionsEdit_Level_Region_CommandTip"
    FormUXHelper.SetTip Edit_Level_Candidate_Command, "ZoneExportGUIOptionsEdit_Level_Candidate_CommandTip"
    FormUXHelper.SetTip GroupBy_Label, "ZoneExportGUIOptionsGroupBy_LabelTip"
    FormUXHelper.SetTip ComboBox_Export_Type, "ZoneExportGUIOptionsGroupBy_LabelTip"
    FormUXHelper.SetTip CheckBox_PerZone, "ZoneExportGUIOptionsPerZone_LabelTip"
    FormUXHelper.SetTip ZoneProperty_Label, "ZoneExportGUIOptionsZoneProperty_LabelTip"
    FormUXHelper.SetTip ComboBox_ZoneProperty, "ZoneExportGUIOptionsZoneProperty_LabelTip"
    FormUXHelper.SetTip Round_Label, "ZoneExportGUIOptionsRound_LabelTip"
    FormUXHelper.SetTip Round_SpinButton, "ZoneExportGUIOptionsRound_LabelTip"
    FormUXHelper.SetTip Use_Dialog_CheckBox, "ZoneExportGUIOptionsUse_Dialog_LabelTip"
    FormUXHelper.SetTip OpenAfter_CheckBox, "ZoneExportGUIOptionsOpenAfter_LabelTip"

    ' Restore-defaults button
    Reset_Command.Caption = GetTranslation("FormResetDefaultsCaption")
    FormUXHelper.SetTip Reset_Command, "FormResetDefaultsTip"

    ' Group-by combo items: localized display, stable English stored key
    ComboBox_Export_Type.Clear
    ComboBox_Export_Type.AddItem GroupByDisplayFromKey("Style")
    ComboBox_Export_Type.AddItem GroupByDisplayFromKey("Level")
    ComboBox_Export_Type.AddItem GroupByDisplayFromKey("Color")
    ComboBox_Export_Type.AddItem GroupByDisplayFromKey("ID")

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
        If TextBox_CandidateLevel.Visible Then FormUXHelper.NudgeActiveEdit TextBox_CandidateLevel
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
    If sGroupBy <> "Level" And sGroupBy <> "Color" And sGroupBy <> "ID" Then sGroupBy = "Style"
    ComboBox_Export_Type.Value = GroupByDisplayFromKey(sGroupBy)

    ' Per-zone split checkbox.
    CheckBox_PerZone.Value = (UCase(Trim(ARESConfig.ARES_ZONE_EXPORT_PER_ZONE.Value)) = "TRUE")

    ' Zone-property combo: populate from the managed custom-property names, then seed the current
    ' value Null-safe (a dropdown-list combo rejects an out-of-list value) - set .Value only when
    ' the stored name is a list member, else .ListIndex = -1 (M1).
    ComboBox_ZoneProperty.Clear
    Dim propNames()    As String
    Dim pi             As Long
    Dim sSelZoneProp   As String
    Dim bZonePropFound As Boolean
    sSelZoneProp = Trim(ARESConfig.ARES_ZONE_EXPORT_ZONE_PROPERTY.Value)
    bZonePropFound = False
    propNames = CustomPropertyHandler.GetCustomPropertyNames()
    For pi = LBound(propNames) To UBound(propNames)
        If Len(Trim(propNames(pi))) > 0 Then
            ComboBox_ZoneProperty.AddItem propNames(pi)
            If StrComp(Trim(propNames(pi)), sSelZoneProp, vbTextCompare) = 0 Then bZonePropFound = True
        End If
    Next pi
    If bZonePropFound Then
        ComboBox_ZoneProperty.Value = sSelZoneProp
    Else
        ComboBox_ZoneProperty.ListIndex = -1
    End If

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
    OpenAfter_CheckBox.Value = (UCase(Trim(ARESConfig.ARES_ZONE_EXPORT_EXCEL_VISIBLE.Value)) = "TRUE")

    TextBox_RegionLevel.Visible = False
    TextBox_CandidateLevel.Visible = False
    ApplyPerZoneEnabled
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.SeedControls"
End Sub

' NOTE: ARES_Zoning_Output_Level is the region output level shown here and is shared with the Zoning form.
Private Sub Reset_Command_Click()
    On Error GoTo ErrorHandler
    FormUXHelper.PersistDefault ARESConfig.ARES_ZONING_OUTPUT_LEVEL
    FormUXHelper.PersistDefault ARESConfig.ARES_ZONE_EXPORT_LEVEL
    FormUXHelper.PersistDefault ARESConfig.ARES_ZONE_EXPORT_GROUP_BY
    FormUXHelper.PersistDefault ARESConfig.ARES_ZONE_EXPORT_PER_ZONE
    FormUXHelper.PersistDefault ARESConfig.ARES_ZONE_EXPORT_ZONE_PROPERTY
    FormUXHelper.PersistDefault ARESConfig.ARES_ZONE_EXPORT_ROUND
    FormUXHelper.PersistDefault ARESConfig.ARES_ZONE_EXPORT_USE_DIALOG
    FormUXHelper.PersistDefault ARESConfig.ARES_ZONE_EXPORT_EXCEL_VISIBLE
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
    ' SetControlsLocked(False) re-enables every combo/checkbox, so re-assert the conditional
    ' zone-property grey-out after each global unlock.
    If Not bState Then ApplyPerZoneEnabled
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.SetLocked"
End Sub

' Conditional grey-out: the zone-property combo + its label are only meaningful when the per-zone
' breakdown is on. Re-asserted after every SetControlsLocked(False) and on every SeedControls. Runs
' only on unlock/seed, when the combo cannot hold focus, so disabling it never ejects focus.
Private Sub ApplyPerZoneEnabled()
    On Error GoTo ErrorHandler
    ' Null-safe read: a triple-state checkbox could yield Null, which errors on a direct Boolean
    ' assign; "If ... = True" treats Null as False (bOn defaults False). Mirrors the IIf idiom used
    ' by the sibling checkbox handlers.
    Dim bOn As Boolean
    If CheckBox_PerZone.Value = True Then bOn = True
    ComboBox_ZoneProperty.Enabled = bOn
    ZoneProperty_Label.Enabled = bOn
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.ApplyPerZoneEnabled"
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
    ElseIf sDisp = GetTranslation("ZoneExportGroupByID") Then
        GroupByKeyFromDisplay = "ID"
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
        Case "ID"
            GroupByDisplayFromKey = GetTranslation("ZoneExportGroupByID")
        Case Else
            GroupByDisplayFromKey = GetTranslation("ZoneExportGroupByStyle")
    End Select
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.GroupByDisplayFromKey"
    GroupByDisplayFromKey = GetTranslation("ZoneExportGroupByStyle")
End Function

