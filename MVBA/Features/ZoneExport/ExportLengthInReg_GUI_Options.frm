' UserForm: ExportLengthInReg_GUI_Options
' Description: Options panel for ExportLengthInRegion — zone level, grouping key, rounding, save dialog.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ARESConfigClass, ErrorHandlerClass
Option Explicit

Private mLocked As Boolean

' ============================================================
' ZONE LEVEL — Edit button + hidden TextBox
' ============================================================

Private Sub Edit_Level_Region_Command_Click()
    On Error GoTo ErrorHandler
    If Not mLocked Then
        Locked
        TextBox_RegionLevel.Value = ARESConfig.ARES_ZONING_OUTPUT_LEVEL.Value
        TextBox_RegionLevel.Visible = True
        Edit_Level_Region_Command.Visible = False
        TextBox_RegionLevel.SetFocus
    End If
    Exit Sub

ErrorHandler:
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
    If mLocked Then Locked
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.TextBox_RegionLevel_Exit"
End Sub

Private Sub TextBox_RegionLevel_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    If Shift = 0 Then
        If KeyCode = 13 Then TextBox_RegionLevel_Exit returnB
        If KeyCode = 27 Then
            TextBox_RegionLevel.Visible = False
            Edit_Level_Region_Command.Visible = True
            If mLocked Then Locked
        End If
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.TextBox_RegionLevel_KeyUp"
End Sub

' ============================================================
' GROUP BY — ComboBox
' ============================================================

Private Sub ComboBox_Export_Type_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = ComboBox_Export_Type.Value
    If sVal <> "Style" And sVal <> "Level" And sVal <> "Color" Then Exit Sub
    If mLocked Then Exit Sub
    If ARESConfig.ARES_ZONE_EXPORT_GROUP_BY.Value <> sVal Then
        Locked
        ARESConfig.ARES_ZONE_EXPORT_GROUP_BY.Value = sVal
        Locked
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.ComboBox_Export_Type_Change"
End Sub

' ============================================================
' ROUNDING — SpinButton
' ============================================================

Private Sub Round_SpinButton_Change()
    On Error GoTo ErrorHandler
    If mLocked Then Exit Sub
    If CStr(Round_SpinButton.Value) <> ARESConfig.ARES_ZONE_EXPORT_ROUND.Value Then
        Locked
        Round_Number_Label.Caption = Round_SpinButton.Value
        ARESConfig.ARES_ZONE_EXPORT_ROUND.Value = CStr(Round_SpinButton.Value)
        Locked
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.Round_SpinButton_Change"
End Sub

' ============================================================
' USE DIALOG — CheckBox
' ============================================================

Private Sub Use_Dialog_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(Use_Dialog_CheckBox.Value, "True", "False")
    If mLocked Then Exit Sub
    If ARESConfig.ARES_ZONE_EXPORT_USE_DIALOG.Value <> sVal Then
        Locked
        ARESConfig.ARES_ZONE_EXPORT_USE_DIALOG.Value = sVal
        Locked
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.Use_Dialog_CheckBox_Change"
End Sub

' ============================================================
' INITIALIZATION
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption = GetTranslation("ZoneExportGUIOptionsCaption")
    Round_Label.Caption = GetTranslation("ZoneExportGUIOptionsRound_LabelCaption")
    Use_Dialog_Label.Caption = GetTranslation("ZoneExportGUIOptionsUse_Dialog_LabelCaption")
    Edit_Level_Region_Command.Caption = GetTranslation("ZoneExportGUIOptionsEdit_Level_Region_CommandCaption")

    ComboBox_Export_Type.Clear
    ComboBox_Export_Type.AddItem "Style"
    ComboBox_Export_Type.AddItem "Level"
    ComboBox_Export_Type.AddItem "Color"
    Dim sGroupBy As String
    sGroupBy = Trim(ARESConfig.ARES_ZONE_EXPORT_GROUP_BY.Value)
    If sGroupBy <> "Level" And sGroupBy <> "Color" Then sGroupBy = "Style"
    ComboBox_Export_Type.Value = sGroupBy

    Round_SpinButton.Min = 0
    Round_SpinButton.Max = 10
    Dim nRound As Integer
    nRound = CInt(ARESConfig.ARES_ZONE_EXPORT_ROUND.Value)
    Round_SpinButton.Value = nRound
    Round_Number_Label.Caption = nRound

    Use_Dialog_CheckBox.Value = (UCase(Trim(ARESConfig.ARES_ZONE_EXPORT_USE_DIALOG.Value)) = "TRUE")

    TextBox_RegionLevel.Visible = False
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.UserForm_Initialize"
End Sub

' ============================================================
' CLOSE
' ============================================================

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHandler
    If mLocked Then
        Cancel = True
        If TextBox_RegionLevel.Visible Then Me.TextBox_RegionLevel.SetFocus
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.UserForm_QueryClose"
End Sub

Private Sub UserForm_Terminate()
    command.OnZoneExportGUIClosed
End Sub

' ============================================================
' LOCK HELPER
' ============================================================

Private Function Locked() As Boolean
    On Error GoTo ErrorHandler
    mLocked = Not mLocked
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        CheckControlForLock ctrl, mLocked
    Next ctrl
    Locked = mLocked
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.Locked"
    Locked = False
End Function

Private Sub CheckControlForLock(ctrl As Control, lockState As Boolean)
    On Error GoTo ErrorHandler
    Select Case TypeName(ctrl)
        Case "CommandButton", "CheckBox", "SpinButton", "ComboBox"
            ctrl.Enabled = Not lockState
        Case "Frame", "MultiPage", "Page"
            Dim subCtrl As Control
            For Each subCtrl In ctrl.Controls
                CheckControlForLock subCtrl, lockState
            Next subCtrl
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ExportLengthInReg_GUI_Options.CheckControlForLock"
End Sub
