' UserForm: AutoLenghts_GUI_Options
' Description: This UserForm is used for editing the option of AutoLenghts
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private mLocked As Boolean
Private MsgBoxOpen As Boolean

Private Sub Cell_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim CVal As String
    
    CVal = IIf(Cell_CheckBox.Value, "True", "False")
    If mLocked Then
    
    ElseIf ARESConfig.ARES_UPDATE_ATLASCELLLABEL.Value <> CVal Then
        Locked
        ARESConfig.ARES_UPDATE_ATLASCELLLABEL.Value = CVal
        Locked
    End If
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLenghts_GUI_Options.Cell_CheckBox_Change"
End Sub

Private Sub Color_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim CVal As String
    
    CVal = IIf(Color_CheckBox.Value, "True", "False")
    If mLocked Then
    
    ElseIf ARESConfig.ARES_UPDATE_COLOR_WITH_LENGTH.Value <> CVal Then
        Locked
        ARESConfig.ARES_UPDATE_COLOR_WITH_LENGTH.Value = CVal
        Locked
    End If
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLenghts_GUI_Options.Color_CheckBox_Change"
End Sub

Private Sub Edit_Cells_List_Command_Click()
    On Error GoTo ErrorHandler
    If Not mLocked Then
        Locked
        TextBox_Cells_List.Value = ARESConfig.ARES_CELL_LIKE_LABEL.Value
        TextBox_Cells_List.Visible = True
        Edit_Cells_List_Command.Visible = False
        TextBox_Cells_List.SetFocus
    End If
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLenghts_GUI_Options.Edit_Cells_List_Command_Click"
End Sub

Private Sub Edit_Trigger_Command_Click()
    On Error GoTo ErrorHandler
    If Not mLocked Then
        Locked
        TextBox_Trigger.Value = ARESConfig.ARES_LENGTH_TRIGGER_ID.Value
        TextBox_Trigger.Visible = True
        Edit_Trigger_Command.Visible = False
        TextBox_Trigger.SetFocus
    End If
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLenghts_GUI_Options.Edit_Trigger_Command_Click"
End Sub

Private Sub Edit_Triggers_List_Command_Click()
    On Error GoTo ErrorHandler
    If Not mLocked Then
        Locked
        TextBox_Triggers_List.Value = ARESConfig.ARES_LENGTH_TRIGGER.Value
        TextBox_Triggers_List.Visible = True
        Edit_Triggers_List_Command.Visible = False
        TextBox_Triggers_List.SetFocus
    End If
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLenghts_GUI_Options.Edit_Triggers_List_Command_Click"
End Sub

Private Sub Main_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim CVal As String
    
    CVal = IIf(Main_CheckBox.Value, "True", "False")
    If mLocked Then
    
    ElseIf ARESConfig.ARES_AUTO_LENGTHS.Value <> CVal Then
        Locked
        ARESConfig.ARES_AUTO_LENGTHS.Value = CVal
        Locked
    End If
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLenghts_GUI_Options.Main_CheckBox_Change"
End Sub

Private Sub Round_SpinButton_Change()
    On Error GoTo ErrorHandler
    If mLocked Then
        
    ElseIf Round_SpinButton.Value <> ARESConfig.ARES_LENGTH_ROUND.Value Then
        Locked
        Round_Number_Label.Caption = Round_SpinButton.Value
        ARESConfig.ARES_LENGTH_ROUND.Value = Round_SpinButton.Value
        Locked
    End If
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLenghts_GUI_Options.Round_SpinButton_Change"
End Sub

Private Sub TextBox_Triggers_List_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    Dim txts() As String
    Dim i As Long
    
    If TextBox_Triggers_List.Value <> ARESConfig.ARES_LENGTH_TRIGGER.Value Then
        txts = Split(TextBox_Triggers_List.Value, ARESConstants.ARES_VAR_DELIMITER)
        For i = LBound(txts) To UBound(txts)
            If Not txts(i) Like "*" & ARESConfig.ARES_LENGTH_TRIGGER_ID.Value & "*" Then
                MsgBoxOpen = True
                MsgBox GetTranslation("AutoLengthsGUIOptionsEdit_Triggers_List_Error") & ARESConfig.ARES_LENGTH_TRIGGER_ID.Value, vbOKOnly
                MsgBoxOpen = False
                Exit Sub
            End If
        Next i
        ARESConfig.ARES_LENGTH_TRIGGER.Value = TextBox_Triggers_List.Value
    End If
    
    TextBox_Triggers_List.Visible = False
    Edit_Triggers_List_Command.Visible = True
    If mLocked Then Locked
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLenghts_GUI_Options.TextBox_Triggers_List_Exit"
End Sub
Private Sub TextBox_Cells_List_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    If TextBox_Cells_List.Value <> ARESConfig.ARES_CELL_LIKE_LABEL.Value Then
        ARESConfig.ARES_CELL_LIKE_LABEL.Value = TextBox_Cells_List.Value
    End If
    TextBox_Cells_List.Visible = False
    Edit_Cells_List_Command.Visible = True
    If mLocked Then Locked
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLenghts_GUI_Options.TextBox_Cells_List_Exit"
End Sub

Private Sub TextBox_Trigger_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    If TextBox_Trigger.Value <> ARESConfig.ARES_LENGTH_TRIGGER_ID.Value Then
        ARESConfig.ARES_LENGTH_TRIGGER.Value = Replace(ARESConfig.ARES_LENGTH_TRIGGER.Value, ARESConfig.ARES_LENGTH_TRIGGER_ID.Value, TextBox_Trigger.Value)
        ARESConfig.ARES_LENGTH_TRIGGER_ID.Value = TextBox_Trigger.Value
    End If
    TextBox_Trigger.Visible = False
    Edit_Trigger_Command.Visible = True
    If mLocked Then Locked
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLenghts_GUI_Options.TextBox_Trigger_Exit"
End Sub

Private Sub TextBox_Triggers_List_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    If Not MsgBoxOpen Then
        If Shift = 0 Then
            If KeyCode = 13 Then
                TextBox_Triggers_List_Exit returnB
            End If
            If KeyCode = 27 Then
                TextBox_Triggers_List.Visible = False
                Edit_Triggers_List_Command.Visible = True
                If mLocked Then Locked
            End If
        End If
    End If
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLenghts_GUI_Options.TextBox_Triggers_List_KeyUp"
End Sub

Private Sub TextBox_Trigger_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    If Shift = 0 Then
        If KeyCode = 13 Then
            TextBox_Trigger_Exit returnB
        End If
        If KeyCode = 27 Then
            TextBox_Trigger.Visible = False
            Edit_Trigger_Command.Visible = True
            If mLocked Then Locked
        End If
    End If
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLenghts_GUI_Options.TextBox_Trigger_KeyUp"
End Sub

Private Sub TextBox_Cells_List_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    If Shift = 0 Then
        If KeyCode = 13 Then
            TextBox_Cells_List_Exit returnB
        End If
        If KeyCode = 27 Then
            TextBox_Cells_List.Visible = False
            Edit_Cells_List_Command.Visible = True
            If mLocked Then Locked
        End If
    End If
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLenghts_GUI_Options.TextBox_Cells_List_KeyUp"
End Sub

' Event handler for initializing the UserForm
' Sets the caption of the form using a translation key
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    
    Me.Caption = GetTranslation("AutoLengthsGUIOptionsCaption")
    Main_Label.Caption = GetTranslation("AutoLengthsGUIOptionsMain_LabelCaption")
    Color_Label.Caption = GetTranslation("AutoLengthsGUIOptionsColor_LabelCaption")
    Cell_Label.Caption = GetTranslation("AutoLengthsGUIOptionsCell_LabelCaption")
    Edit_Trigger_Command.Caption = GetTranslation("AutoLengthsGUIOptionsEdit_Trigger_CommandCaption", ARESConfig.ARES_LENGTH_TRIGGER_ID.Value)
    Edit_Triggers_List_Command.Caption = GetTranslation("AutoLengthsGUIOptionsEdit_Triggers_List_CommandCaption")
    Round_Label.Caption = GetTranslation("AutoLengthsGUIOptionsRound_LabelCaption")
    Edit_Cells_List_Command.Caption = GetTranslation("AutoLengthsGUIOptionsEdit_Cells_List_CommandCaption")
    Round_Number_Label.Caption = ARESConfig.ARES_LENGTH_ROUND.Value
    Round_SpinButton.Value = Round_Number_Label.Caption
    If ARESConfig.ARES_AUTO_LENGTHS.Value Then
        Main_CheckBox.Value = "True"
    Else
        Main_CheckBox.Value = "False"
    End If
    If ARESConfig.ARES_UPDATE_COLOR_WITH_LENGTH.Value Then
        Color_CheckBox.Value = "True"
    Else
        Color_CheckBox.Value = "False"
    End If
    If ARESConfig.ARES_UPDATE_ATLASCELLLABEL.Value Then
        Cell_CheckBox.Value = "True"
    Else
        Cell_CheckBox.Value = "False"
    End If
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLenghts_GUI_Options.UserForm_Initialize"
End Sub
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
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLenghts_GUI_Options.Locked"
    Locked = False
End Function

Private Sub CheckControlForLock(ctrl As Control, lockState As Boolean)
    On Error GoTo ErrorHandler
    If TypeName(ctrl) = "CommandButton" Or TypeName(ctrl) = "CheckBox" Or TypeName(ctrl) = "SpinButton" Then
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
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLenghts_GUI_Options.CheckControlForLock"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHandler
    If mLocked Then
        Cancel = True
        Select Case True
            Case TextBox_Triggers_List.Visible
                Me.TextBox_Triggers_List.SetFocus
                SeeActiveTextBox TextBox_Triggers_List
            Case TextBox_Cells_List.Visible
                Me.TextBox_Cells_List.SetFocus
                SeeActiveTextBox TextBox_Cells_List
            Case TextBox_Trigger.Visible
                Me.TextBox_Trigger.SetFocus
                SeeActiveTextBox TextBox_Trigger
        End Select
    End If
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLenghts_GUI_Options.UserForm_QueryClose"
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
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLenghts_GUI_Options.SeeActiveTextBox"
End Sub
