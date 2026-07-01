' UserForm: PropertyTagging_GUI_Options
' Description: UserForm for editing custom-property (Property Tagging) options.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ErrorHandlerClass, ARESConfigClass, PropertyTagging
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private mbLocked As Boolean

' ============================================================
' MASTER SWITCH — CheckBox -> ARES_Auto_Properties
' ============================================================

Private Sub Main_CheckBox_Change()
    On Error GoTo ErrorHandler
    Dim CVal As String
    CVal = IIf(Main_CheckBox.Value, "True", "False")
    If mbLocked Then
    ElseIf ARESConfig.ARES_AUTO_PROPERTIES.Value <> CVal Then
        Locked
        ARESConfig.ARES_AUTO_PROPERTIES.Value = CVal
        Locked
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.Main_CheckBox_Change"
End Sub

' ============================================================
' CUSTOM PROPERTY LIST — Edit button + hidden TextBox -> ARES_Custom_Property_List
' ============================================================

Private Sub Edit_PropertyList_Command_Click()
    On Error GoTo ErrorHandler
    If Not mbLocked Then
        Locked
        TextBox_PropertyList.Value = ARESConfig.ARES_CUSTOM_PROPERTY_LIST.Value
        TextBox_PropertyList.Visible = True
        Edit_PropertyList_Command.Visible = False
        TextBox_PropertyList.SetFocus
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.Edit_PropertyList_Command_Click"
End Sub

Private Sub TextBox_PropertyList_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    If TextBox_PropertyList.Value <> ARESConfig.ARES_CUSTOM_PROPERTY_LIST.Value Then
        ARESConfig.ARES_CUSTOM_PROPERTY_LIST.Value = TextBox_PropertyList.Value
    End If
    TextBox_PropertyList.Visible = False
    Edit_PropertyList_Command.Visible = True
    If mbLocked Then Locked
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.TextBox_PropertyList_Exit"
End Sub

Private Sub TextBox_PropertyList_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    If Shift = 0 Then
        If KeyCode = 13 Then TextBox_PropertyList_Exit returnB
        If KeyCode = 27 Then
            TextBox_PropertyList.Visible = False
            Edit_PropertyList_Command.Visible = True
            If mbLocked Then Locked
        End If
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.TextBox_PropertyList_KeyUp"
End Sub

' ============================================================
' PROPERTY RULES — Edit button + hidden TextBox -> ARES_Property_Rules
' On write, refresh PropertyTagging's parsed cache so the new rules take effect immediately.
' ============================================================

Private Sub Edit_Rules_Command_Click()
    On Error GoTo ErrorHandler
    If Not mbLocked Then
        Locked
        TextBox_Rules.Value = ARESConfig.ARES_PROPERTY_RULES.Value
        TextBox_Rules.Visible = True
        Edit_Rules_Command.Visible = False
        TextBox_Rules.SetFocus
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.Edit_Rules_Command_Click"
End Sub

Private Sub TextBox_Rules_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    If TextBox_Rules.Value <> ARESConfig.ARES_PROPERTY_RULES.Value Then
        ARESConfig.ARES_PROPERTY_RULES.Value = TextBox_Rules.Value
        PropertyTagging.RefreshRules            ' apply the edited rules live, no restart
    End If
    TextBox_Rules.Visible = False
    Edit_Rules_Command.Visible = True
    If mbLocked Then Locked
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.TextBox_Rules_Exit"
End Sub

Private Sub TextBox_Rules_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Dim returnB As MSForms.ReturnBoolean
    If Shift = 0 Then
        If KeyCode = 13 Then TextBox_Rules_Exit returnB
        If KeyCode = 27 Then
            TextBox_Rules.Visible = False
            Edit_Rules_Command.Visible = True
            If mbLocked Then Locked
        End If
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.TextBox_Rules_KeyUp"
End Sub

' ============================================================
' FORM LIFECYCLE
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption                    = GetTranslation("PropertyTaggingGUIOptionsCaption")
    Main_Label.Caption            = GetTranslation("PropertyTaggingGUIOptionsMain_LabelCaption")
    Edit_PropertyList_Command.Caption = GetTranslation("PropertyTaggingGUIOptionsEditList_CommandCaption")
    Edit_Rules_Command.Caption    = GetTranslation("PropertyTaggingGUIOptionsEditRules_CommandCaption")

    If ARESConfig.ARES_AUTO_PROPERTIES.Value Then
        Main_CheckBox.Value = "True"
    Else
        Main_CheckBox.Value = "False"
    End If

    TextBox_PropertyList.Visible = False
    TextBox_Rules.Visible        = False
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.UserForm_Initialize"
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
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.Locked"
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
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.CheckControlForLock"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHandler

    If mbLocked Then
        Cancel = True
        Select Case True
            Case TextBox_PropertyList.Visible
                Me.TextBox_PropertyList.SetFocus
                SeeActiveTextBox TextBox_PropertyList
            Case TextBox_Rules.Visible
                Me.TextBox_Rules.SetFocus
                SeeActiveTextBox TextBox_Rules
        End Select
    Else
        command.OnPropertyTaggingGUIClosed
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.UserForm_QueryClose"
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
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyTagging_GUI_Options.SeeActiveTextBox"
End Sub
