VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Outline_GUI_Options 
   Caption         =   "Edit zoning options:"
   ClientHeight    =   3285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3015
   OleObjectBlob   =   "Outline_GUI_Options.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Outline_GUI_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm: Outline_GUI_Options
' Description: UserForm for editing Outline module options (tight per-element zoning variant).
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
        TextBox_Levels.Value = ARESConfig.ARES_OUTLINE_LEVEL.Value
        TextBox_Levels.Visible = True
        Edit_Levels_Command.Visible = False
        TextBox_Levels.SetFocus
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Outline_GUI_Options.Edit_Levels_Command_Click"
End Sub

Private Sub TextBox_Levels_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    FormUXHelper.CommitInlineEdit TextBox_Levels, Edit_Levels_Command, ARESConfig.ARES_OUTLINE_LEVEL
    SetLocked False
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Outline_GUI_Options.TextBox_Levels_Exit"
End Sub

Private Sub TextBox_Levels_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            TextBox_Levels_Exit returnB
        Case FormUXKeyCancel
            FormUXHelper.RevertInlineEdit TextBox_Levels, ARESConfig.ARES_OUTLINE_LEVEL
            TextBox_Levels_Exit returnB
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Outline_GUI_Options.TextBox_Levels_KeyUp"
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
        LangManager.ShowStatusT "OutlineGUIOptionsDistanceError"
        TextBox_Distance.Value = ARESConfig.ARES_OUTLINE_DISTANCE.Value
        Exit Sub
    End If
    If sNorm <> ARESConfig.ARES_OUTLINE_DISTANCE.Value Then
        ARESConfig.ARES_OUTLINE_DISTANCE.Value = sNorm
        TextBox_Distance.Value = sNorm
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Outline_GUI_Options.TextBox_Distance_Exit"
End Sub

' ============================================================
' OUTPUT LEVEL - Edit button + hidden TextBox
' ============================================================

Private Sub Edit_OutputLevel_Command_Click()
    On Error GoTo ErrorHandler
    If Not mbLocked Then
        SetLocked True
        TextBox_OutputLevel.Value = ARESConfig.ARES_OUTLINE_OUTPUT_LEVEL.Value
        TextBox_OutputLevel.Visible = True
        Edit_OutputLevel_Command.Visible = False
        TextBox_OutputLevel.SetFocus
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Outline_GUI_Options.Edit_OutputLevel_Command_Click"
End Sub

Private Sub TextBox_OutputLevel_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    FormUXHelper.CommitInlineEdit TextBox_OutputLevel, Edit_OutputLevel_Command, ARESConfig.ARES_OUTLINE_OUTPUT_LEVEL
    Edit_OutputLevel_Command.Caption = GetTranslation("OutlineGUIOptionsEditOutputLevel_CommandCaption", ARESConfig.ARES_OUTLINE_OUTPUT_LEVEL.Value)
    SetLocked False
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Outline_GUI_Options.TextBox_OutputLevel_Exit"
End Sub

Private Sub TextBox_OutputLevel_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            TextBox_OutputLevel_Exit returnB
        Case FormUXKeyCancel
            FormUXHelper.RevertInlineEdit TextBox_OutputLevel, ARESConfig.ARES_OUTLINE_OUTPUT_LEVEL
            TextBox_OutputLevel_Exit returnB
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Outline_GUI_Options.TextBox_OutputLevel_KeyUp"
End Sub

' ============================================================
' OUTPUT COLOR - CommandButton with BackColor preview
' ============================================================

Private Sub Edit_Color_Command_Click()
    On Error GoTo ErrorHandler
    If mbLocked Then Exit Sub

    Dim newIdx As Long
    newIdx = ColorDialog.PickMsColorIndex(CLng(ARESConfig.ARES_OUTLINE_OUTPUT_COLOR.Value))
    If newIdx = -1 Then Exit Sub

    ColorDialog.ApplyColorToTextBox TextBox_Color, newIdx
    ARESConfig.ARES_OUTLINE_OUTPUT_COLOR.Value = CStr(newIdx)
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Outline_GUI_Options.Edit_Color_Command_Click"
End Sub

' ============================================================
' OUTPUT STYLE - always-visible TextBox (supports custom style names)
' ============================================================

Private Sub TextBox_Style_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    If TextBox_Style.Value <> ARESConfig.ARES_OUTLINE_OUTPUT_STYLE.Value Then
        ARESConfig.ARES_OUTLINE_OUTPUT_STYLE.Value = TextBox_Style.Value
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Outline_GUI_Options.TextBox_Style_Exit"
End Sub

' ============================================================
' OUTPUT WEIGHT - SpinButton (0-31)
' ============================================================

Private Sub Weight_SpinButton_Change()
    On Error GoTo ErrorHandler
    If Not mbLocked And Weight_SpinButton.Value <> CLng(ARESConfig.ARES_OUTLINE_OUTPUT_WEIGHT.Value) Then
        SetLocked True
        Weight_Number_Label.Caption = Weight_SpinButton.Value
        ARESConfig.ARES_OUTLINE_OUTPUT_WEIGHT.Value = CStr(Weight_SpinButton.Value)
        SetLocked False
    End If
    Exit Sub

ErrorHandler:
    SetLocked False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Outline_GUI_Options.Weight_SpinButton_Change"
End Sub

' ============================================================
' FORM LIFECYCLE
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption = GetTranslation("OutlineGUIOptionsCaption")
    Edit_Levels_Command.Caption = GetTranslation("OutlineGUIOptionsEditLevels_CommandCaption")
    Distance_Label.Caption = GetTranslation("OutlineGUIOptionsDistance_LabelCaption")
    Edit_OutputLevel_Command.Caption = GetTranslation("OutlineGUIOptionsEditOutputLevel_CommandCaption", ARESConfig.ARES_OUTLINE_OUTPUT_LEVEL.Value)
    Edit_OutputLevel_Command.WordWrap = True
    Edit_Color_Command.Caption = GetTranslation("OutlineGUIOptionsEditColor_CommandCaption")
    Style_Label.Caption = GetTranslation("OutlineGUIOptionsOutputStyle_LabelCaption")
    Weight_Label.Caption = GetTranslation("OutlineGUIOptionsWeight_LabelCaption")

    ' Tooltips (AC-6)
    FormUXHelper.SetTip Edit_Levels_Command, "OutlineGUIOptionsEditLevels_CommandTip"
    FormUXHelper.SetTip Distance_Label, "OutlineGUIOptionsDistance_LabelTip"
    FormUXHelper.SetTip TextBox_Distance, "OutlineGUIOptionsDistance_LabelTip"
    FormUXHelper.SetTip Edit_OutputLevel_Command, "OutlineGUIOptionsEditOutputLevel_CommandTip"
    FormUXHelper.SetTip Edit_Color_Command, "OutlineGUIOptionsEditColor_CommandTip"
    FormUXHelper.SetTip TextBox_Color, "OutlineGUIOptionsColor_SwatchTip"
    FormUXHelper.SetTip Style_Label, "OutlineGUIOptionsOutputStyle_LabelTip"
    FormUXHelper.SetTip TextBox_Style, "OutlineGUIOptionsOutputStyle_LabelTip"
    FormUXHelper.SetTip Weight_Label, "OutlineGUIOptionsWeight_LabelTip"
    FormUXHelper.SetTip Weight_SpinButton, "OutlineGUIOptionsWeight_LabelTip"

    ' Keyboard order + mnemonics (AC-7) - existing controls only
    Edit_Levels_Command.TabIndex = 0
    TextBox_Distance.TabIndex = 1
    Edit_OutputLevel_Command.TabIndex = 2
    Edit_Color_Command.TabIndex = 3
    TextBox_Style.TabIndex = 4
    Weight_SpinButton.TabIndex = 5
    Edit_Levels_Command.Accelerator = "L"
    Edit_OutputLevel_Command.Accelerator = "O"
    Edit_Color_Command.Accelerator = "C"

    Dim storedDist As String
    storedDist = Replace(ARESConfig.ARES_OUTLINE_DISTANCE.Value, ",", ".")
    If Val(storedDist) <= 0 Then
        storedDist = ARESConfig.ARES_OUTLINE_DISTANCE.DefaultValue
        ARESConfig.ARES_OUTLINE_DISTANCE.Value = storedDist
    End If
    TextBox_Distance.Value = storedDist

    TextBox_Color.Locked = True
    ColorDialog.ApplyColorToTextBox TextBox_Color, CLng(ARESConfig.ARES_OUTLINE_OUTPUT_COLOR.Value)

    TextBox_Style.Value = ARESConfig.ARES_OUTLINE_OUTPUT_STYLE.Value

    Weight_SpinButton.Min = 0:  Weight_SpinButton.Max = 31

    Weight_Number_Label.Caption = ARESConfig.ARES_OUTLINE_OUTPUT_WEIGHT.Value
    Weight_SpinButton.Value = CLng(ARESConfig.ARES_OUTLINE_OUTPUT_WEIGHT.Value)

    TextBox_Levels.Visible = False
    TextBox_OutputLevel.Visible = False
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Outline_GUI_Options.UserForm_Initialize"
End Sub

' Explicit-state lock (AC-2/AC-8): replaces the toggle Locked()/CheckControlForLock pair.
' Any error path must call SetLocked False so controls are never left disabled.
Private Sub SetLocked(ByVal bState As Boolean)
    On Error GoTo ErrorHandler
    mbLocked = bState
    FormUXHelper.SetControlsLocked Me, bState
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Outline_GUI_Options.SetLocked"
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
        command.OnOutlineGUIClosed
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Outline_GUI_Options.UserForm_QueryClose"
End Sub

