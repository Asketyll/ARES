VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConfigThemes_GUI 
   Caption         =   "ConfigThemes_GUI"
   ClientHeight    =   2655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3270
   OleObjectBlob   =   "ConfigThemes_GUI.frx":0000
   StartUpPosition =   0  'Manual
End
Attribute VB_Name = "ConfigThemes_GUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm: ConfigThemes_GUI
' Description: The theme window (key-in OpenARESThemes): picking a theme in the list loads it at once, and the
'              current settings can be saved as a theme under a name.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ConfigThemes, LangManager, ErrorHandlerClass, ARESConfigClass, FormUXHelper, FormPlacement, Command
Option Explicit

' True while the controls are seeded: a selection made by the seed is not a pick and must not load a theme.
Private mbSeeding As Boolean
' True between the two DropButtonClick calls of one drop-down (list shown, then hidden).
Private mbListOpen As Boolean

' ============================================================
' THEME LIST - picking an entry loads that theme (arrow keys included)
' ============================================================

Private Sub ComboBox_Theme_Change()
    On Error GoTo ErrorHandler
    If mbSeeding Then Exit Sub
    If ComboBox_Theme.ListIndex < 0 Then Exit Sub
    ConfigThemes.LoadTheme CStr(ComboBox_Theme.List(ComboBox_Theme.ListIndex))
    ' A refused load selects the displayed theme again
    SeedControls
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes_GUI.ComboBox_Theme_Change"
End Sub

' MSForms raises DropButtonClick when the list appears AND when it disappears. As the list appears it is re-listed
' (a theme added in Explorer shows up) and the selection is cleared, so picking the displayed theme is a change
' that reloads it. A list closed without a pick shows the displayed theme again.
Private Sub ComboBox_Theme_DropButtonClick()
    On Error GoTo ErrorHandler
    If mbSeeding Then Exit Sub
    mbSeeding = True
    If Not mbListOpen Then
        mbListOpen = True
        FillThemeList
        If ComboBox_Theme.ListIndex <> -1 Then ComboBox_Theme.ListIndex = -1
    Else
        mbListOpen = False
        If ComboBox_Theme.ListIndex < 0 Then SelectTheme ARESConfig.ARES_THEME_CURRENT.value
    End If
    mbSeeding = False
    Exit Sub

ErrorHandler:
    mbSeeding = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes_GUI.ComboBox_Theme_DropButtonClick"
End Sub

' ============================================================
' SAVE - name box + Save button (Enter in the box saves too)
' ============================================================

Private Sub SaveTheme_Command_Click()
    On Error GoTo ErrorHandler
    SaveTypedTheme
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes_GUI.SaveTheme_Command_Click"
End Sub

Private Sub TextBox_Name_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    FormUXHelper.NoteInlineKeyDown KeyCode, Shift
    ' Swallowed: MSForms would otherwise move the focus to the next control, and the key-up would land there.
    If KeyCode = vbKeyReturn And Shift = 0 Then KeyCode = 0
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes_GUI.TextBox_Name_KeyDown"
End Sub

Private Sub TextBox_Name_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error GoTo ErrorHandler
    Select Case FormUXHelper.InlineEditKey(KeyCode, Shift)
        Case FormUXKeyCommit
            SaveTypedTheme
        Case FormUXKeyCancel
            TextBox_Name.value = ARESConfig.ARES_THEME_CURRENT.value
    End Select
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes_GUI.TextBox_Name_KeyUp"
End Sub

' A refused or cancelled save keeps the typed name in the box, to be corrected.
Private Sub SaveTypedTheme()
    On Error GoTo ErrorHandler
    SeedControls bKeepName:=Not ConfigThemes.SaveTheme(CStr(TextBox_Name.value))
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes_GUI.SaveTypedTheme"
End Sub

' ============================================================
' FORM LIFECYCLE
' ============================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Me.Caption = GetTranslation("ConfigThemesGUICaption")
    Theme_Label.Caption = GetTranslation("ConfigThemesGUITheme_LabelCaption")
    Name_Label.Caption = GetTranslation("ConfigThemesGUIName_LabelCaption")
    SaveTheme_Command.Caption = GetTranslation("ConfigThemesGUISaveTheme_CommandCaption")

    ' Tooltips
    FormUXHelper.SetTip Theme_Label, "ConfigThemesGUITheme_Tip"
    FormUXHelper.SetTip ComboBox_Theme, "ConfigThemesGUITheme_Tip"
    Name_Label.ControlTipText = GetTranslation("ConfigThemesGUIName_Tip", ConfigThemes.ThemeFolder())
    TextBox_Name.ControlTipText = Name_Label.ControlTipText
    FormUXHelper.SetTip SaveTheme_Command, "ConfigThemesGUISaveTheme_Tip"

    SeedControls
    FormPlacement.RestoreFormPosition Me, Me.Name
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes_GUI.UserForm_Initialize"
End Sub

' Called after every configuration load (Command.RefreshOpenForms).
Public Sub RefreshFromConfig()
    On Error GoTo ErrorHandler
    SeedControls
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes_GUI.RefreshFromConfig"
End Sub

' List the themes, select the displayed one (none when it is unset or its file is gone), and prefill the name
' box with it unless bKeepName.
Private Sub SeedControls(Optional ByVal bKeepName As Boolean = False)
    On Error GoTo ErrorHandler
    Dim bWasSeeding As Boolean
    bWasSeeding = mbSeeding
    mbSeeding = True

    FillThemeList
    SelectTheme ARESConfig.ARES_THEME_CURRENT.value
    If Not bKeepName Then TextBox_Name.value = ARESConfig.ARES_THEME_CURRENT.value

    mbSeeding = bWasSeeding
    Exit Sub

ErrorHandler:
    mbSeeding = bWasSeeding
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes_GUI.SeedControls"
End Sub

' Fill the list from the theme folder; returns True when the list had to be rebuilt. An empty folder shows the
' no-theme line and disables the list.
Private Function FillThemeList() As Boolean
    On Error GoTo ErrorHandler
    Dim sNames() As String
    Dim nThemes As Long
    Dim i As Long
    Dim bSame As Boolean

    FillThemeList = False
    nThemes = ConfigThemes.ListThemes(sNames)

    bSame = (ComboBox_Theme.ListCount = nThemes)
    If bSame Then
        For i = 0 To nThemes - 1
            If CStr(ComboBox_Theme.List(i)) <> sNames(i) Then
                bSame = False
                Exit For
            End If
        Next i
    End If

    If Not bSame Then
        ComboBox_Theme.Clear
        For i = 0 To nThemes - 1
            ComboBox_Theme.AddItem sNames(i)
        Next i
        FillThemeList = True
    End If

    NoTheme_Label.Caption = GetTranslation("ConfigThemeNoTheme", ConfigThemes.ThemeFolder())
    NoTheme_Label.Visible = (nThemes = 0)
    ComboBox_Theme.Enabled = (nThemes > 0)
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes_GUI.FillThemeList"
End Function

Private Sub SelectTheme(ByVal sName As String)
    On Error GoTo ErrorHandler
    Dim i As Long
    For i = 0 To ComboBox_Theme.ListCount - 1
        If StrComp(CStr(ComboBox_Theme.List(i)), sName, vbTextCompare) = 0 Then
            If ComboBox_Theme.ListIndex <> i Then ComboBox_Theme.ListIndex = i
            Exit Sub
        End If
    Next i
    If ComboBox_Theme.ListIndex <> -1 Then ComboBox_Theme.ListIndex = -1
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes_GUI.SelectTheme"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHandler
    FormPlacement.SaveFormPosition Me, Me.Name
    Command.OnConfigThemesGUIClosed
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes_GUI.UserForm_QueryClose"
End Sub
