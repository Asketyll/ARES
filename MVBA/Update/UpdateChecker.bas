' Module: UpdateChecker
' Description: Checks for available ARES updates via the GitHub Releases API.
'              The installed version is written to the Windows Registry by the installer.
'              User notification preferences are stored as MicroStation configuration variables.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (via BootLoader)
Option Explicit

Private Const ARES_REGISTRY_KEY As String = "HKCU\Software\ARES\Version"
Private Const ARES_GITHUB_API_URL As String = "https://api.github.com/repos/Asketyll/ARES/releases/latest"

' Returns the installed version from the Windows Registry.
' Falls back to ARES_CONFIG_VERSION if the key is absent, empty, or on any error.
Public Function GetInstalledVersion() As String
    On Error Resume Next

    Dim oShell As Object
    Dim sVersion As String

    Set oShell = CreateObject("WScript.Shell")
    sVersion = oShell.RegRead(ARES_REGISTRY_KEY)
    Set oShell = Nothing

    If Err.Number <> 0 Or Len(Trim(sVersion)) = 0 Then
        Err.Clear
        GetInstalledVersion = ARES_CONFIG_VERSION
    Else
        GetInstalledVersion = sVersion
    End If
End Function

' Returns the latest release tag from the GitHub API (without leading "v").
' Returns an empty string on network failure, HTTP error, or parse error — never raises.
Public Function GetLatestVersionFromGitHub() As String
    On Error GoTo ErrorHandler

    Dim oHttp As Object
    Dim sResponse As String
    Dim sTagName As String
    Dim lStart As Long
    Dim lEnd As Long

    Set oHttp = CreateObject("MSXML2.XMLHTTP")
    oHttp.Open "GET", ARES_GITHUB_API_URL, False
    oHttp.setRequestHeader "User-Agent", "ARES-MVBA"
    oHttp.send

    If oHttp.Status <> 200 Then GoTo ErrorHandler

    sResponse = oHttp.responseText
    Set oHttp = Nothing

    ' Extract tag_name value from JSON response
    lStart = InStr(sResponse, """tag_name""")
    If lStart = 0 Then GoTo ErrorHandler

    lStart = InStr(lStart, sResponse, ":")
    If lStart = 0 Then GoTo ErrorHandler

    lStart = InStr(lStart, sResponse, """")
    If lStart = 0 Then GoTo ErrorHandler
    lStart = lStart + 1

    lEnd = InStr(lStart, sResponse, """")
    If lEnd = 0 Then GoTo ErrorHandler

    sTagName = Mid(sResponse, lStart, lEnd - lStart)

    ' Strip leading "v" if present (e.g. "v1.2.3" -> "1.2.3")
    If Left(sTagName, 1) = "v" Then sTagName = Mid(sTagName, 2)

    GetLatestVersionFromGitHub = sTagName
    Exit Function

ErrorHandler:
    GetLatestVersionFromGitHub = ""
End Function

' Checks for an available update and notifies the user if one is found.
' Silently exits if the network is unavailable, the check fails, or preferences suppress it.
Public Sub CheckForUpdate()
    On Error Resume Next

    Dim sInstalled As String
    Dim sLatest As String

    ' [1] Check if notifications are permanently muted
    If ARESConfig.ARES_UPDATE_MUTE.Value = "True" Then Exit Sub

    ' [2] Get installed version (registry, fallback to ARES_CONFIG_VERSION constant)
    sInstalled = GetInstalledVersion()

    ' [3] Get latest version from GitHub — empty string means network unavailable
    sLatest = GetLatestVersionFromGitHub()
    If Len(sLatest) = 0 Then Exit Sub

    ' [4] Already up to date
    If sInstalled = sLatest Then Exit Sub

    ' [5] User previously chose to ignore this specific version
    If ARESConfig.ARES_UPDATE_IGNORE_VERSION.Value = sLatest Then Exit Sub

    ' [6] Show notification dialog
    ShowUpdateDialog sInstalled, sLatest
End Sub

' Shows the update notification dialog and processes the user's choice.
' Uses vbYesNoCancel: Yes=Ignore this version | No=Never remind me | Cancel=Remind later
Private Sub ShowUpdateDialog(ByVal sInstalled As String, ByVal sLatest As String)
    On Error Resume Next

    Dim sMsg As String
    Dim nResult As Integer

    sMsg = "A new version of ARES is available." & vbCrLf & vbCrLf & _
           "Installed : " & sInstalled & vbCrLf & _
           "Available : " & sLatest & vbCrLf & vbCrLf & _
           "https://github.com/Asketyll/ARES/releases/latest" & vbCrLf & vbCrLf & _
           "[Yes]    Ignore this version" & vbCrLf & _
           "[No]     Never remind me" & vbCrLf & _
           "[Cancel] Remind me later"

    nResult = MsgBox(sMsg, vbYesNoCancel + vbInformation, "ARES - Update Available")

    Select Case nResult
        Case vbYes    ' Ignore this version only
            ARESConfig.ARES_UPDATE_IGNORE_VERSION.Value = sLatest
        Case vbNo     ' Disable all future notifications
            ARESConfig.ARES_UPDATE_MUTE.Value = "True"
        Case vbCancel ' Remind me later — no state change
    End Select
End Sub
