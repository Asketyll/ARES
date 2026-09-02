VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CableReport_GUI_Options 
   Caption         =   "CableReport_GUI_Options"
   ClientHeight    =   5790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "CableReport_GUI_Options.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CableReport_GUI_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm: CableReport_GUI_Options
' Description: Options panel for CableReport - cable level, zone level, the 4 custom-property
'              names (Repere/Nature/Longueur/soil-type), end-cell search radius, rounding,
'              Excel-visible toggle. Mirrors ExportLengthInReg_GUI_Options's structure/UX baseline.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ARESConfigClass, ErrorHandlerClass, FormUXHelper, FormPlacement, CustomPropertyHandler
Option Explicit

Private mbLocked As Boolean

' ============================================================
' CABLE LEVEL - Edit button + hidden TextBox (required: the field cannot be blanked here,
' mirrors ExportLengthInReg's Region Level - blanking a required level is a footgun to avoid
' in the GUI, even though the runtime already handles an empty value safely).
' ============================================================

Private Sub Edit_CableLevel_Command_Click()
    On Error GoTo ErrorHandler
    If Not mbLocked Then
        SetLocked True
        TextBox_CableLevel.value = ARESConfig.ARES_CABLEREPORT_CABLE_LEVEL.value
        TextBox_CableLevel.Visible = True
        Edit_CableLevel_Command.Visible = False
        TextBox_CableLevel.SetFocus
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.Edit_CableLevel_Command_Click"
End Sub

Private Sub TextBox_CableLevel_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = Trim(TextBox_CableLevel.value)
    If Len(sVal) > 0 And sVal <> ARESConfig.ARES_CABLEREPORT_CABLE_LEVEL.value Then
        ARESConfig.ARES_CABLEREPORT_CABLE_LEVEL.value = sVal
    End If
    TextBox_CableLevel.Visible = False
    Edit_CableLevel_Command.Caption = GetTranslation("CableReportGUIOptionsEditCableLevel_CommandCaption")
    Edit_CableLevel_Command.Visible = True
    SetLocked False
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.TextBox_CableLevel_Exit"
End Sub

Private Sub TextBox_CableLevel_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    FormUXHelper.NoteInlineKeyDown KeyCode, Shift
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.TextBox_CableLevel_KeyDown"
End Sub

Private Sub TextBox_CableLevel_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            TextBox_CableLevel_Exit returnB
            Edit_CableLevel_Command.SetFocus
        Case FormUXKeyCancel
            FormUXHelper.RevertInlineEdit TextBox_CableLevel, ARESConfig.ARES_CABLEREPORT_CABLE_LEVEL
            TextBox_CableLevel_Exit returnB
            Edit_CableLevel_Command.SetFocus
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.TextBox_CableLevel_KeyUp"
End Sub

' ============================================================
' ZONE LEVEL - Edit button + hidden TextBox (optional: empty is valid, falls back to
' ARES_Outline_Output_Level at runtime - mirrors ExportLengthInReg's Candidate Level, no
' Len>0 guard on write).
' ============================================================

Private Sub Edit_ZoneLevel_Command_Click()
    On Error GoTo ErrorHandler
    If Not mbLocked Then
        SetLocked True
        TextBox_ZoneLevel.value = ARESConfig.ARES_CABLEREPORT_ZONE_LEVEL.value
        TextBox_ZoneLevel.Visible = True
        Edit_ZoneLevel_Command.Visible = False
        TextBox_ZoneLevel.SetFocus
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.Edit_ZoneLevel_Command_Click"
End Sub

Private Sub TextBox_ZoneLevel_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    ' Empty is a valid value here (falls back to ARES_Outline_Output_Level), so no Len>0 guard.
    Dim sVal As String
    sVal = Trim(TextBox_ZoneLevel.value)
    If sVal <> ARESConfig.ARES_CABLEREPORT_ZONE_LEVEL.value Then
        ARESConfig.ARES_CABLEREPORT_ZONE_LEVEL.value = sVal
    End If
    TextBox_ZoneLevel.Visible = False
    Edit_ZoneLevel_Command.Caption = GetTranslation("CableReportGUIOptionsEditZoneLevel_CommandCaption")
    Edit_ZoneLevel_Command.Visible = True
    SetLocked False
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.TextBox_ZoneLevel_Exit"
End Sub

Private Sub TextBox_ZoneLevel_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    FormUXHelper.NoteInlineKeyDown KeyCode, Shift
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.TextBox_ZoneLevel_KeyDown"
End Sub

Private Sub TextBox_ZoneLevel_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            TextBox_ZoneLevel_Exit returnB
            Edit_ZoneLevel_Command.SetFocus
        Case FormUXKeyCancel
            FormUXHelper.RevertInlineEdit TextBox_ZoneLevel, ARESConfig.ARES_CABLEREPORT_ZONE_LEVEL
            TextBox_ZoneLevel_Exit returnB
            Edit_ZoneLevel_Command.SetFocus
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.TextBox_ZoneLevel_KeyUp"
End Sub

' ============================================================
' CUSTOM-PROPERTY NAME COMBOS (Repere / Nature / Longueur / soil-type) - all 4 populated from
' the ARES DGNLib's ItemTypes (CustomPropertyHandler.GetCustomPropertyNames), same Null-safe
' read/seed pattern as ExportLengthInReg's zone-property combo. Shared logic factored into
' SeedPropertyCombo/CommitPropertyCombo below since all 4 are otherwise identical.
' ============================================================

Private Sub ComboBox_RepereProperty_Change()
    On Error GoTo ErrorHandler
    CommitPropertyCombo ComboBox_RepereProperty, ARESConfig.ARES_CABLEREPORT_REPERE_PROPERTY
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.ComboBox_RepereProperty_Change"
End Sub

Private Sub ComboBox_NatureProperty_Change()
    On Error GoTo ErrorHandler
    CommitPropertyCombo ComboBox_NatureProperty, ARESConfig.ARES_CABLEREPORT_NATURE_PROPERTY
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.ComboBox_NatureProperty_Change"
End Sub

Private Sub ComboBox_LongueurProperty_Change()
    On Error GoTo ErrorHandler
    CommitPropertyCombo ComboBox_LongueurProperty, ARESConfig.ARES_CABLEREPORT_LONGUEUR_PROPERTY
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.ComboBox_LongueurProperty_Change"
End Sub

Private Sub ComboBox_ZoneProperty_Change()
    On Error GoTo ErrorHandler
    CommitPropertyCombo ComboBox_ZoneProperty, ARESConfig.ARES_CABLEREPORT_ZONE_PROPERTY
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.ComboBox_ZoneProperty_Change"
End Sub

' ============================================================
' SEARCH RADIUS - always-visible TextBox (short numeric value, no reveal/hide needed)
' ============================================================

Private Sub TextBox_SearchRadius_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    Dim dVal As Double
    dVal = Val(Trim(Replace(TextBox_SearchRadius.value, ",", ".")))
    If dVal <= 0 Then
        ShowStatusT "CableReportGUIOptionsSearchRadiusError"
        TextBox_SearchRadius.value = ARESConfig.ARES_CABLEREPORT_SEARCH_RADIUS.value
        Exit Sub
    End If
    ' Locale-independent numeric string (dot separator), mirrors FormPlacement.FmtPct's own
    ' Replace(Format(...), ",", ".") idiom - CStr/Format alone follow the Windows locale, and a
    ' French-locale comma written here would silently truncate on the next Val() read.
    Dim sVal As String
    sVal = Replace(Format(dVal, "0.####"), ",", ".")
    If ARESConfig.ARES_CABLEREPORT_SEARCH_RADIUS.value <> sVal Then
        ARESConfig.ARES_CABLEREPORT_SEARCH_RADIUS.value = sVal
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.TextBox_SearchRadius_Exit"
End Sub

' Enter commits in place. Without this the key is not handled here at all: MSForms passes it on to the
' next control (it lands on the Excel-visible CheckBox and toggles it), and the typed radius is only
' written if the user happens to leave the box afterwards. KeyCode = 0 swallows the key so it cannot
' reach that control. The reveal-style level boxes need no such handler - they run the shared
' FormUXHelper arming dance, which an always-visible box does not require.
Private Sub TextBox_SearchRadius_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    If KeyCode = vbKeyReturn And Shift = 0 Then
        TextBox_SearchRadius_Exit returnB
        KeyCode = 0
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.TextBox_SearchRadius_KeyDown"
End Sub

' ============================================================
' ROUNDING - SpinButton (0-10, identical pattern to ExportLengthInReg's Round_SpinButton)
' ============================================================

Private Sub Round_SpinButton_Change()
    On Error GoTo ErrorHandler
    If Not mbLocked And CStr(Round_SpinButton.value) <> ARESConfig.ARES_CABLEREPORT_ROUND.value Then
        SetLocked True
        Round_Number_Label.Caption = Round_SpinButton.value
        ARESConfig.ARES_CABLEREPORT_ROUND.value = CStr(Round_SpinButton.value)
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.Round_SpinButton_Change"
End Sub

' ============================================================
' OPEN AFTER EXPORT - CheckBox (surfaces ARES_CableReport_Excel_Visible)
' ============================================================

Private Sub OpenAfter_CheckBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    If Shift = 0 And KeyCode = vbKeyReturn Then OpenAfter_CheckBox.value = Not OpenAfter_CheckBox.value
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.OpenAfter_CheckBox_KeyUp"
End Sub

Private Sub OpenAfter_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim sVal As String
    sVal = IIf(OpenAfter_CheckBox.value, "True", "False")
    If Not mbLocked And ARESConfig.ARES_CABLEREPORT_EXCEL_VISIBLE.value <> sVal Then
        SetLocked True
        ARESConfig.ARES_CABLEREPORT_EXCEL_VISIBLE.value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.OpenAfter_CheckBox_Change"
End Sub

' ============================================================
' INITIALIZATION
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption = GetTranslation("CableReportGUIOptionsCaption")
    Edit_CableLevel_Command.Caption = GetTranslation("CableReportGUIOptionsEditCableLevel_CommandCaption")
    Edit_ZoneLevel_Command.Caption = GetTranslation("CableReportGUIOptionsEditZoneLevel_CommandCaption")
    RepereProperty_Label.Caption = GetTranslation("CableReportGUIOptionsRepereProperty_LabelCaption")
    NatureProperty_Label.Caption = GetTranslation("CableReportGUIOptionsNatureProperty_LabelCaption")
    LongueurProperty_Label.Caption = GetTranslation("CableReportGUIOptionsLongueurProperty_LabelCaption")
    ZoneProperty_Label.Caption = GetTranslation("CableReportGUIOptionsZoneProperty_LabelCaption")
    SearchRadius_Label.Caption = GetTranslation("CableReportGUIOptionsSearchRadius_LabelCaption")
    Round_Label.Caption = GetTranslation("CableReportGUIOptionsRound_LabelCaption")
    OpenAfter_CheckBox.Caption = GetTranslation("CableReportGUIOptionsOpenAfter_LabelCaption")

    ' Tooltips
    FormUXHelper.SetTip Edit_CableLevel_Command, "CableReportGUIOptionsEditCableLevel_CommandTip"
    FormUXHelper.SetTip Edit_ZoneLevel_Command, "CableReportGUIOptionsEditZoneLevel_CommandTip"
    FormUXHelper.SetTip RepereProperty_Label, "CableReportGUIOptionsRepereProperty_LabelTip"
    FormUXHelper.SetTip ComboBox_RepereProperty, "CableReportGUIOptionsRepereProperty_LabelTip"
    FormUXHelper.SetTip NatureProperty_Label, "CableReportGUIOptionsNatureProperty_LabelTip"
    FormUXHelper.SetTip ComboBox_NatureProperty, "CableReportGUIOptionsNatureProperty_LabelTip"
    FormUXHelper.SetTip LongueurProperty_Label, "CableReportGUIOptionsLongueurProperty_LabelTip"
    FormUXHelper.SetTip ComboBox_LongueurProperty, "CableReportGUIOptionsLongueurProperty_LabelTip"
    FormUXHelper.SetTip ZoneProperty_Label, "CableReportGUIOptionsZoneProperty_LabelTip"
    FormUXHelper.SetTip ComboBox_ZoneProperty, "CableReportGUIOptionsZoneProperty_LabelTip"
    FormUXHelper.SetTip SearchRadius_Label, "CableReportGUIOptionsSearchRadius_LabelTip"
    FormUXHelper.SetTip TextBox_SearchRadius, "CableReportGUIOptionsSearchRadius_LabelTip"
    FormUXHelper.SetTip Round_Label, "CableReportGUIOptionsRound_LabelTip"
    FormUXHelper.SetTip Round_SpinButton, "CableReportGUIOptionsRound_LabelTip"
    FormUXHelper.SetTip OpenAfter_CheckBox, "CableReportGUIOptionsOpenAfter_LabelTip"

    ' Restore-defaults button
    Reset_Command.Caption = GetTranslation("FormResetDefaultsCaption")
    FormUXHelper.SetTip Reset_Command, "FormResetDefaultsTip"

    ' Rounding spin bounds (value seeded in SeedControls, guarded against non-numeric config)
    Round_SpinButton.Min = 0
    Round_SpinButton.Max = 10

    SeedControls
    FormPlacement.RestoreFormPosition Me, Me.Name
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.UserForm_Initialize"
End Sub

' ============================================================
' CLOSE
' ============================================================

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    If mbLocked Then
        Cancel = True
        If TextBox_CableLevel.Visible Then FormUXHelper.NudgeActiveEdit TextBox_CableLevel
        If TextBox_ZoneLevel.Visible Then FormUXHelper.NudgeActiveEdit TextBox_ZoneLevel
    Else
        ' The radius box writes through on _Exit, which does NOT fire when the form is closed with the
        ' caret still in it - commit here so a typed value is never silently lost. Idempotent: the
        ' write itself is compare-guarded. The level boxes need no equivalent; mbLocked blocks the
        ' close while one of them is open for editing.
        TextBox_SearchRadius_Exit returnB
        FormPlacement.SaveFormPosition Me, Me.Name
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.UserForm_QueryClose"
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    command.OnCableReportGUIClosed
End Sub

' ============================================================
' HELPERS
' ============================================================

' Re-seed all controls from the current config values.
Private Sub SeedControls()
    On Error GoTo ErrorHandler

    TextBox_CableLevel.Visible = False
    TextBox_ZoneLevel.Visible = False

    SeedPropertyCombo ComboBox_RepereProperty, ARESConfig.ARES_CABLEREPORT_REPERE_PROPERTY
    SeedPropertyCombo ComboBox_NatureProperty, ARESConfig.ARES_CABLEREPORT_NATURE_PROPERTY
    SeedPropertyCombo ComboBox_LongueurProperty, ARESConfig.ARES_CABLEREPORT_LONGUEUR_PROPERTY
    SeedPropertyCombo ComboBox_ZoneProperty, ARESConfig.ARES_CABLEREPORT_ZONE_PROPERTY

    Dim sRadius As String
    sRadius = ARESConfig.ARES_CABLEREPORT_SEARCH_RADIUS.value
    If Not IsNumeric(sRadius) Then sRadius = ARESConfig.ARES_CABLEREPORT_SEARCH_RADIUS.DefaultValue
    TextBox_SearchRadius.value = sRadius

    Dim nRound As Integer
    If IsNumeric(ARESConfig.ARES_CABLEREPORT_ROUND.value) Then
        nRound = CInt(ARESConfig.ARES_CABLEREPORT_ROUND.value)
    Else
        nRound = CInt(ARESConfig.ARES_CABLEREPORT_ROUND.DefaultValue)
    End If
    If nRound < 0 Then nRound = 0
    If nRound > 10 Then nRound = 10
    Round_SpinButton.value = nRound
    Round_Number_Label.Caption = nRound

    OpenAfter_CheckBox.value = (UCase(Trim(ARESConfig.ARES_CABLEREPORT_EXCEL_VISIBLE.value)) = "TRUE")
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.SeedControls"
End Sub

Private Sub Reset_Command_Click()
    On Error GoTo ErrorHandler
    If Not FormUXHelper.ConfirmReset() Then Exit Sub
    FormUXHelper.PersistDefault ARESConfig.ARES_CABLEREPORT_CABLE_LEVEL
    FormUXHelper.PersistDefault ARESConfig.ARES_CABLEREPORT_ZONE_LEVEL
    FormUXHelper.PersistDefault ARESConfig.ARES_CABLEREPORT_REPERE_PROPERTY
    FormUXHelper.PersistDefault ARESConfig.ARES_CABLEREPORT_NATURE_PROPERTY
    FormUXHelper.PersistDefault ARESConfig.ARES_CABLEREPORT_LONGUEUR_PROPERTY
    FormUXHelper.PersistDefault ARESConfig.ARES_CABLEREPORT_ZONE_PROPERTY
    FormUXHelper.PersistDefault ARESConfig.ARES_CABLEREPORT_SEARCH_RADIUS
    FormUXHelper.PersistDefault ARESConfig.ARES_CABLEREPORT_ROUND
    FormUXHelper.PersistDefault ARESConfig.ARES_CABLEREPORT_EXCEL_VISIBLE
    SeedControls
    LangManager.ShowStatusT "FormDefaultsRestored"
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.Reset_Command_Click"
End Sub

' Explicit-state lock: replaces the toggle Locked()/CheckControlForLock pair.
Private Sub SetLocked(ByVal bState As Boolean)
    On Error GoTo ErrorHandler
    mbLocked = bState
    FormUXHelper.SetControlsLocked Me, bState
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.SetLocked"
End Sub

' Shared by the 4 property-name combos: populate from the ARES DGNLib's ItemTypes, then seed
' Null-safe from oVar's current value (list member -> select it; not a member -> ListIndex = -1,
' same M1 fix as ExportLengthInReg's zone-property combo).
Private Sub SeedPropertyCombo(ByVal oCombo As MSForms.ComboBox, ByVal oVar As ARES_MS_VAR_Class)
    On Error GoTo ErrorHandler
    oCombo.Clear
    Dim propNames() As String
    Dim pi          As Long
    Dim sSel        As String
    Dim bFound      As Boolean
    sSel = Trim(oVar.value)
    bFound = False
    propNames = CustomPropertyHandler.GetCustomPropertyNames()
    For pi = LBound(propNames) To UBound(propNames)
        If Len(Trim(propNames(pi))) > 0 Then
            oCombo.AddItem propNames(pi)
            If StrComp(Trim(propNames(pi)), sSel, vbTextCompare) = 0 Then bFound = True
        End If
    Next pi
    If bFound Then
        oCombo.value = sSel
    Else
        oCombo.ListIndex = -1
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.SeedPropertyCombo"
End Sub

' Shared by the 4 property-name combos: write-through a Null-safe combo value into oVar, only
' on a real change (a dropdown-list combo with no selection returns Null).
Private Sub CommitPropertyCombo(ByVal oCombo As MSForms.ComboBox, ByVal oVar As ARES_MS_VAR_Class)
    On Error GoTo ErrorHandler
    If mbLocked Then Exit Sub
    Dim sVal As String
    If IsNull(oCombo.value) Then
        sVal = ""
    Else
        sVal = Trim(CStr(oCombo.value))
    End If
    If oVar.value <> sVal Then
        SetLocked True
        oVar.value = sVal
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CableReport_GUI_Options.CommitPropertyCombo"
End Sub

