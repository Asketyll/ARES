' UserForm: Zoning_GUI_Options
' Description: UserForm for editing Zoning module options.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass, ARESConfigClass, ColorDialog
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private mbLocked As Boolean

' ============================================================
' SOURCE LEVELS — Edit button + hidden TextBox
' ============================================================

Private Sub Edit_Levels_Command_Click()
    On Error GoTo ErrorHandler
    If Not mbLocked Then
        Locked
        TextBox_Levels.Value = ARESConfig.ARES_ZONING_LEVEL.Value
        TextBox_Levels.Visible = True
        Edit_Levels_Command.Visible = False
        TextBox_Levels.SetFocus
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.Edit_Levels_Command_Click"
End Sub

Private Sub TextBox_Levels_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    If TextBox_Levels.Value <> ARESConfig.ARES_ZONING_LEVEL.Value Then
        ARESConfig.ARES_ZONING_LEVEL.Value = TextBox_Levels.Value
    End If
    TextBox_Levels.Visible = False
    Edit_Levels_Command.Visible = True
    If mbLocked Then Locked
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.TextBox_Levels_Exit"
End Sub

Private Sub TextBox_Levels_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    If Shift = 0 Then
        If KeyCode = 13 Then TextBox_Levels_Exit returnB
        If KeyCode = 27 Then
            TextBox_Levels.Visible = False
            Edit_Levels_Command.Visible = True
            If mbLocked Then Locked
        End If
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.TextBox_Levels_KeyUp"
End Sub

' ============================================================
' BUFFER DISTANCE — always-visible TextBox
' ============================================================

Private Sub TextBox_Distance_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    Dim normalized As String
    normalized = Replace(TextBox_Distance.Value, ",", ".")
    If Val(normalized) <= 0 Then
        Cancel = True
        MsgBox GetTranslation("ZoningGUIOptionsDistanceError"), vbOKOnly
        TextBox_Distance.Value = ARESConfig.ARES_ZONING_DISTANCE.Value
        Exit Sub
    End If
    If normalized <> ARESConfig.ARES_ZONING_DISTANCE.Value Then
        ARESConfig.ARES_ZONING_DISTANCE.Value = normalized
        TextBox_Distance.Value = normalized
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.TextBox_Distance_Exit"
End Sub

' ============================================================
' OUTPUT LEVEL — Edit button + hidden TextBox
' ============================================================

Private Sub Edit_OutputLevel_Command_Click()
    On Error GoTo ErrorHandler
    If Not mbLocked Then
        Locked
        TextBox_OutputLevel.Value = ARESConfig.ARES_ZONING_OUTPUT_LEVEL.Value
        TextBox_OutputLevel.Visible = True
        Edit_OutputLevel_Command.Visible = False
        TextBox_OutputLevel.SetFocus
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.Edit_OutputLevel_Command_Click"
End Sub

Private Sub TextBox_OutputLevel_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    If TextBox_OutputLevel.Value <> ARESConfig.ARES_ZONING_OUTPUT_LEVEL.Value Then
        ARESConfig.ARES_ZONING_OUTPUT_LEVEL.Value = TextBox_OutputLevel.Value
    End If
    Edit_OutputLevel_Command.Caption = GetTranslation("ZoningGUIOptionsEditOutputLevel_CommandCaption", ARESConfig.ARES_ZONING_OUTPUT_LEVEL.Value)
    TextBox_OutputLevel.Visible = False
    Edit_OutputLevel_Command.Visible = True
    If mbLocked Then Locked
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.TextBox_OutputLevel_Exit"
End Sub

Private Sub TextBox_OutputLevel_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    If Shift = 0 Then
        If KeyCode = 13 Then TextBox_OutputLevel_Exit returnB
        If KeyCode = 27 Then
            TextBox_OutputLevel.Visible = False
            Edit_OutputLevel_Command.Visible = True
            If mbLocked Then Locked
        End If
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.TextBox_OutputLevel_KeyUp"
End Sub

' ============================================================
' OUTPUT COLOR — CommandButton with BackColor preview
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
' OUTPUT STYLE — always-visible TextBox (supports custom style names)
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
' OUTPUT WEIGHT — SpinButton (0-31)
' ============================================================

Private Sub Weight_SpinButton_Change()
    On Error GoTo ErrorHandler
    If mbLocked Then
    ElseIf Weight_SpinButton.Value <> CLng(ARESConfig.ARES_ZONING_OUTPUT_WEIGHT.Value) Then
        Locked
        Weight_Number_Label.Caption = Weight_SpinButton.Value
        ARESConfig.ARES_ZONING_OUTPUT_WEIGHT.Value = CStr(Weight_SpinButton.Value)
        Locked
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.Weight_SpinButton_Change"
End Sub

' ============================================================
' FORM LIFECYCLE
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption                       = GetTranslation("ZoningGUIOptionsCaption")
    Edit_Levels_Command.Caption      = GetTranslation("ZoningGUIOptionsEditLevels_CommandCaption")
    Distance_Label.Caption           = GetTranslation("ZoningGUIOptionsDistance_LabelCaption")
    Edit_OutputLevel_Command.Caption  = GetTranslation("ZoningGUIOptionsEditOutputLevel_CommandCaption", ARESConfig.ARES_ZONING_OUTPUT_LEVEL.Value)
    Edit_OutputLevel_Command.WordWrap = True
    Edit_Color_Command.Caption       = GetTranslation("ZoningGUIOptionsEditColor_CommandCaption")
    Style_Label.Caption              = GetTranslation("ZoningGUIOptionsOutputStyle_LabelCaption")
    Weight_Label.Caption             = GetTranslation("ZoningGUIOptionsWeight_LabelCaption")

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
    Weight_SpinButton.Value     = CLng(ARESConfig.ARES_ZONING_OUTPUT_WEIGHT.Value)

    TextBox_Levels.Visible      = False
    TextBox_OutputLevel.Visible = False
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.UserForm_Initialize"
End Sub

Private Function Locked() As Boolean
    On Error GoTo ErrorHandler
    mbLocked = Not mbLocked

    Dim ctrl As Control
    For Each ctrl In Me.Controls
        CheckControlForLock ctrl, mbLocked
    Next ctrl

    Locked = mbLocked
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.Locked"
    Locked = False
End Function

Private Sub CheckControlForLock(ctrl As Control, lockState As Boolean)
    On Error GoTo ErrorHandler
    If TypeName(ctrl) = "CommandButton" Or TypeName(ctrl) = "SpinButton" Then
        ctrl.Enabled = Not lockState
    Else
        If TypeName(ctrl) = "Frame" Or TypeName(ctrl) = "MultiPage" Or TypeName(ctrl) = "Page" Then
            Dim subCtrl As Control
            For Each subCtrl In ctrl.Controls
                CheckControlForLock subCtrl, lockState
            Next subCtrl
        End If
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.CheckControlForLock"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHandler

    If mbLocked Then
        Cancel = True
        Select Case True
            Case TextBox_Levels.Visible
                Me.TextBox_Levels.SetFocus
                SeeActiveTextBox TextBox_Levels
            Case TextBox_OutputLevel.Visible
                Me.TextBox_OutputLevel.SetFocus
                SeeActiveTextBox TextBox_OutputLevel
        End Select
    Else
        command.OnZoningGUIClosed
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.UserForm_QueryClose"
End Sub

Private Sub SeeActiveTextBox(ctrl As TextBox)
    On Error GoTo ErrorHandler
    Dim i As Byte
    For i = 0 To 3
        ctrl.SpecialEffect = fmSpecialEffectBump
        DoEvents
        Sleep 75
        ctrl.SpecialEffect = fmSpecialEffectSunken
        DoEvents
        Sleep 75
    Next i
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Zoning_GUI_Options.SeeActiveTextBox"
End Sub
