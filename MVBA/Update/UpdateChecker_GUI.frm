VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UpdateChecker_GUI
   Caption         =   "New update available"
   ClientHeight    =   1230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4350
   OleObjectBlob   =   "UpdateChecker_GUI.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UpdateChecker_GUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm: UpdateChecker_GUI
' Description: Update notification dialog - asks user whether to update, skip this version, or mute all.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: LangManager, ARESConfigClass (via BootLoader), FormUXHelper
Option Explicit

' Captions/tooltips are set once in Initialize (was Activate under On Error Resume Next).
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    Me.Caption = GetTranslation("UpdateAvailableTitle")
    lblQuestion.Caption = GetTranslation("UpdateAvailableQuestion")
    cmdYes.Caption = GetTranslation("UpdateBtnYes")
    cmdNo.Caption = GetTranslation("UpdateBtnSkipVersion")     ' truthful label (was "No"): skips THIS version only
    cmdIgnoreAll.Caption = GetTranslation("UpdateBtnIgnoreAll")

    ' Tooltips (AC-6/AC-11)
    FormUXHelper.SetTip cmdYes, "UpdateBtnYesTip"
    FormUXHelper.SetTip cmdNo, "UpdateBtnSkipVersionTip"
    FormUXHelper.SetTip cmdIgnoreAll, "UpdateBtnIgnoreAllTip"

    ' Guaranteed non-empty fallback captions if translations are not ready.
    If Len(Me.Caption) = 0 Then Me.Caption = "ARES - Update Available"
    If Len(cmdYes.Caption) = 0 Then cmdYes.Caption = "Yes"
    If Len(cmdNo.Caption) = 0 Then cmdNo.Caption = "Skip this version"
    If Len(cmdIgnoreAll.Caption) = 0 Then cmdIgnoreAll.Caption = "Ignore all"
    Exit Sub

ErrorHandler:
    If Len(Me.Caption) = 0 Then Me.Caption = "ARES - Update Available"
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "UpdateChecker_GUI.UserForm_Initialize"
End Sub

Private Sub cmdYes_Click()
    On Error GoTo ErrorHandler
    Me.Hide
    DownloadAndInstall
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "UpdateChecker_GUI.cmdYes_Click"
End Sub

' "Skip this version": permanently skip the current release (writes ARES_Update_Ignore_Version).
Private Sub cmdNo_Click()
    On Error GoTo ErrorHandler
    ARESConfig.ARES_UPDATE_IGNORE_VERSION.Value = GetUpdateLatestVersion()
    Me.Hide
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "UpdateChecker_GUI.cmdNo_Click"
End Sub

' "Ignore all": mute ALL future update prompts (writes ARES_Update_Mute).
Private Sub cmdIgnoreAll_Click()
    On Error GoTo ErrorHandler
    ARESConfig.ARES_UPDATE_MUTE.Value = "True"
    Me.Hide
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "UpdateChecker_GUI.cmdIgnoreAll_Click"
End Sub

' Intended "remind me later" affordance: closing via the window [X] writes NO config, so ARES simply
' offers the update again on the next launch. An explicit "Later" button is Track B (story 8-1, section 6.1).
