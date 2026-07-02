VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Zoning_GUI_Options 
   Caption         =   "Edit zoning options:"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3015
   OleObjectBlob   =   "Zoning_GUI_Options.frx":0000
   StartUpPosition =   0  'Manual
End
Attribute VB_Name = "Zoning_GUI_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm: Zoning_GUI_Options
' Description: UserForm for editing Zoning module options.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass, ARESConfigClass, ColorDialog
Option Explicit

Private mbLocked As Boolean

' ============================================================
' SOURCE LEVELS - Edit button + hidden TextBox
' ============================================================

Private Sub Edit_Levels_Command_Click()
    On Error GoTo ErrorHandler
    If Not mbLocked Then
        SetLocked True
        TextBox_Levels.Value = ARESConfig.ARES_ZONING_LEVEL.Value
        TextBox_Levels.Visible = True
        Edit_Levels_Command.Visible = False
        TextBox_Levels.SetFocus
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.Edit_Levels_Command_Click"
End Sub

Private Sub TextBox_Levels_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    FormUXHelper.CommitInlineEdit TextBox_Levels, Edit_Levels_Command, ARESConfig.ARES_ZONING_LEVEL
    SetLocked False
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.TextBox_Levels_Exit"
End Sub

Private Sub TextBox_Levels_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    FormUXHelper.NoteInlineKeyDown KeyCode, Shift
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.TextBox_Levels_KeyDown"
End Sub

Private Sub TextBox_Levels_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            TextBox_Levels_Exit returnB
            Edit_Levels_Command.SetFocus
        Case FormUXKeyCancel
            FormUXHelper.RevertInlineEdit TextBox_Levels, ARESConfig.ARES_ZONING_LEVEL
            TextBox_Levels_Exit returnB
            Edit_Levels_Command.SetFocus
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.TextBox_Levels_KeyUp"
End Sub

' ============================================================
' BUFFER DISTANCE - always-visible TextBox
' ============================================================

Private Sub TextBox_Distance_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    Dim sNorm As String
    sNorm = Replace(TextBox_Distance.Value, ",", ".")
    If Val(sNorm) <= 0 Then
        Cancel = True
        LangManager.ShowStatusT "ZoningGUIOptionsDistanceError"
        TextBox_Distance.Value = ARESConfig.ARES_ZONING_DISTANCE.Value
        Exit Sub
    End If
    If sNorm <> ARESConfig.ARES_ZONING_DISTANCE.Value Then
        ARESConfig.ARES_ZONING_DISTANCE.Value = sNorm
        TextBox_Distance.Value = sNorm
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.TextBox_Distance_Exit"
End Sub

' ============================================================
' OUTPUT LEVEL - Edit button + hidden TextBox
' ============================================================

Private Sub Edit_OutputLevel_Command_Click()
    On Error GoTo ErrorHandler
    If Not mbLocked Then
        SetLocked True
        TextBox_OutputLevel.Value = ARESConfig.ARES_ZONING_OUTPUT_LEVEL.Value
        TextBox_OutputLevel.Visible = True
        Edit_OutputLevel_Command.Visible = False
        TextBox_OutputLevel.SetFocus
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.Edit_OutputLevel_Command_Click"
End Sub

Private Sub TextBox_OutputLevel_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    FormUXHelper.CommitInlineEdit TextBox_OutputLevel, Edit_OutputLevel_Command, ARESConfig.ARES_ZONING_OUTPUT_LEVEL
    Edit_OutputLevel_Command.Caption = GetTranslation("ZoningGUIOptionsEditOutputLevel_CommandCaption", ARESConfig.ARES_ZONING_OUTPUT_LEVEL.Value)
    SetLocked False
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.TextBox_OutputLevel_Exit"
End Sub

Private Sub TextBox_OutputLevel_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    FormUXHelper.NoteInlineKeyDown KeyCode, Shift
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.TextBox_OutputLevel_KeyDown"
End Sub

Private Sub TextBox_OutputLevel_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            TextBox_OutputLevel_Exit returnB
            Edit_OutputLevel_Command.SetFocus
        Case FormUXKeyCancel
            FormUXHelper.RevertInlineEdit TextBox_OutputLevel, ARESConfig.ARES_ZONING_OUTPUT_LEVEL
            TextBox_OutputLevel_Exit returnB
            Edit_OutputLevel_Command.SetFocus
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.TextBox_OutputLevel_KeyUp"
End Sub

' ============================================================
' OUTPUT COLOR - CommandButton with BackColor preview
' ============================================================

Private Sub Edit_Color_Command_Click()
    On Error GoTo ErrorHandler
    If mbLocked Then Exit Sub

    Dim newIdx As Long
    newIdx = ColorDialog.PickMsColorIndex(CLng(ARESConfig.ARES_ZONING_OUTPUT_COLOR.Value))
    If newIdx = -1 Then Exit Sub

    ColorDialog.ApplyColorToTextBox TextBox_Color, newIdx
    ARESConfig.ARES_ZONING_OUTPUT_COLOR.Value = CStr(newIdx)
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.Edit_Color_Command_Click"
End Sub

' ============================================================
' OUTPUT STYLE - always-visible TextBox (supports custom style names)
' ============================================================

Private Sub TextBox_Style_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    If TextBox_Style.Value <> ARESConfig.ARES_ZONING_OUTPUT_STYLE.Value Then
        ARESConfig.ARES_ZONING_OUTPUT_STYLE.Value = TextBox_Style.Value
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.TextBox_Style_Exit"
End Sub

' ============================================================
' OUTPUT WEIGHT - SpinButton (0-31)
' ============================================================

Private Sub Weight_SpinButton_Change()
    On Error GoTo ErrorHandler
    If Not mbLocked And Weight_SpinButton.Value <> CLng(ARESConfig.ARES_ZONING_OUTPUT_WEIGHT.Value) Then
        SetLocked True
        Weight_Number_Label.Caption = Weight_SpinButton.Value
        ARESConfig.ARES_ZONING_OUTPUT_WEIGHT.Value = CStr(Weight_SpinButton.Value)
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.Weight_SpinButton_Change"
End Sub

' ============================================================
' FORM LIFECYCLE
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption = GetTranslation("ZoningGUIOptionsCaption")
    Edit_Levels_Command.Caption = GetTranslation("ZoningGUIOptionsEditLevels_CommandCaption")
    Distance_Label.Caption = GetTranslation("ZoningGUIOptionsDistance_LabelCaption")
    Edit_OutputLevel_Command.WordWrap = True
    Edit_Color_Command.Caption = GetTranslation("ZoningGUIOptionsEditColor_CommandCaption")
    Style_Label.Caption = GetTranslation("ZoningGUIOptionsOutputStyle_LabelCaption")
    Weight_Label.Caption = GetTranslation("ZoningGUIOptionsWeight_LabelCaption")

    ' Tooltips (AC-6)
    FormUXHelper.SetTip Edit_Levels_Command, "ZoningGUIOptionsEditLevels_CommandTip"
    FormUXHelper.SetTip Distance_Label, "ZoningGUIOptionsDistance_LabelTip"
    FormUXHelper.SetTip TextBox_Distance, "ZoningGUIOptionsDistance_LabelTip"
    FormUXHelper.SetTip Edit_OutputLevel_Command, "ZoningGUIOptionsEditOutputLevel_CommandTip"
    FormUXHelper.SetTip Edit_Color_Command, "ZoningGUIOptionsEditColor_CommandTip"
    FormUXHelper.SetTip TextBox_Color, "ZoningGUIOptionsColor_SwatchTip"
    FormUXHelper.SetTip Style_Label, "ZoningGUIOptionsOutputStyle_LabelTip"
    FormUXHelper.SetTip TextBox_Style, "ZoningGUIOptionsOutputStyle_LabelTip"
    FormUXHelper.SetTip Weight_Label, "ZoningGUIOptionsWeight_LabelTip"
    FormUXHelper.SetTip Weight_SpinButton, "ZoningGUIOptionsWeight_LabelTip"


    ' Restore-defaults button
    Reset_Command.Caption = GetTranslation("FormResetDefaultsCaption")
    FormUXHelper.SetTip Reset_Command, "FormResetDefaultsTip"

    SeedControls
    FormPlacement.RestoreFormPosition Me, Me.Name
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.UserForm_Initialize"
End Sub

' Re-seed all controls from the current config values.
Private Sub SeedControls()
    On Error GoTo ErrorHandler
    Edit_OutputLevel_Command.Caption = GetTranslation("ZoningGUIOptionsEditOutputLevel_CommandCaption", ARESConfig.ARES_ZONING_OUTPUT_LEVEL.Value)

    Dim storedDist As String
    storedDist = Replace(ARESConfig.ARES_ZONING_DISTANCE.Value, ",", ".")
    If Val(storedDist) <= 0 Then
        storedDist = ARESConfig.ARES_ZONING_DISTANCE.DefaultValue
        ARESConfig.ARES_ZONING_DISTANCE.Value = storedDist
    End If
    TextBox_Distance.Value = storedDist

    TextBox_Color.Locked = True
    ColorDialog.ApplyColorToTextBox TextBox_Color, CLng(ARESConfig.ARES_ZONING_OUTPUT_COLOR.Value)

    TextBox_Style.Value = ARESConfig.ARES_ZONING_OUTPUT_STYLE.Value

    Weight_SpinButton.Min = 0:  Weight_SpinButton.Max = 31
    Weight_Number_Label.Caption = ARESConfig.ARES_ZONING_OUTPUT_WEIGHT.Value
    Weight_SpinButton.Value = CLng(ARESConfig.ARES_ZONING_OUTPUT_WEIGHT.Value)

    TextBox_Levels.Visible = False
    TextBox_OutputLevel.Visible = False
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.SeedControls"
End Sub

' Restore every option this form edits to its default value, persist, then re-seed.
Private Sub Reset_Command_Click()
    On Error GoTo ErrorHandler
    FormUXHelper.PersistDefault ARESConfig.ARES_ZONING_LEVEL
    FormUXHelper.PersistDefault ARESConfig.ARES_ZONING_DISTANCE
    FormUXHelper.PersistDefault ARESConfig.ARES_ZONING_OUTPUT_LEVEL
    FormUXHelper.PersistDefault ARESConfig.ARES_ZONING_OUTPUT_COLOR
    FormUXHelper.PersistDefault ARESConfig.ARES_ZONING_OUTPUT_STYLE
    FormUXHelper.PersistDefault ARESConfig.ARES_ZONING_OUTPUT_WEIGHT
    SeedControls
    LangManager.ShowStatusT "FormDefaultsRestored"
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.Reset_Command_Click"
End Sub

' Any error path must call SetLocked False so controls are never left disabled.
Private Sub SetLocked(ByVal bState As Boolean)
    On Error GoTo ErrorHandler
    mbLocked = bState
    FormUXHelper.SetControlsLocked Me, bState
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.SetLocked"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHandler

    If mbLocked Then
        Cancel = True
        Select Case True
            Case TextBox_Levels.Visible
                FormUXHelper.NudgeActiveEdit TextBox_Levels
            Case TextBox_OutputLevel.Visible
                FormUXHelper.NudgeActiveEdit TextBox_OutputLevel
        End Select
    Else
        FormPlacement.SaveFormPosition Me, Me.Name
        command.OnZoningGUIClosed
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.UserForm_QueryClose"
End Sub
