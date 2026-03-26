' UserForm: UpdateChecker_GUI
' Description: Update notification dialog — asks user whether to update, skip, or mute.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ARESConfigClass (via BootLoader)

Option Explicit

Private Sub UserForm_Activate()
    On Error Resume Next
    Me.Caption = GetTranslation("UpdateAvailableTitle")
    lblQuestion.Caption = GetTranslation("UpdateAvailableQuestion")
    cmdYes.Caption = GetTranslation("UpdateBtnYes")
    cmdNo.Caption = GetTranslation("UpdateBtnNo")
    cmdIgnoreAll.Caption = GetTranslation("UpdateBtnIgnoreAll")
End Sub

Private Sub cmdYes_Click()
    Me.Hide
    DownloadAndInstall
End Sub

Private Sub cmdNo_Click()
    On Error Resume Next
    ARESConfig.ARES_UPDATE_IGNORE_VERSION.Value = gsUpdateLatestVersion
    Me.Hide
End Sub

Private Sub cmdIgnoreAll_Click()
    On Error Resume Next
    ARESConfig.ARES_UPDATE_MUTE.Value = "True"
    Me.Hide
End Sub
