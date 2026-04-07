' Module: UpdateChecker
' Description: Checks for available ARES updates via the GitHub Releases API.
'              The installed version is written to the Windows Registry by the installer.
'              User notification preferences are stored as MicroStation configuration variables.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (via BootLoader)
Option Explicit

Private Const ARES_REGISTRY_KEY As String = "HKCU\Software\ARES\Version"
Private Const ARES_GITHUB_API_URL As String = "https://api.github.com/repos/Asketyll/ARES/releases/latest"
Private Const ARES_GITHUB_DOWNLOAD_URL As String = "https://github.com/Asketyll/ARES/releases/download/v{0}/ARES.mvba"
Private Const ARES_MVBA_PATH As String = "C:\ARES\ARES.mvba"

' Module-level vars read by UpdateChecker_GUI
Public gsUpdateLatestVersion As String
Public gsUpdateDownloadUrl As String

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

' Shows the update notification dialog (UpdateChecker_GUI).
' Buttons: Update (open browser) | Ignore this version | Ignore all updates
Private Sub ShowUpdateDialog(ByVal sInstalled As String, ByVal sLatest As String)
    On Error Resume Next

    gsUpdateLatestVersion = sLatest
    gsUpdateDownloadUrl = Replace(ARES_GITHUB_DOWNLOAD_URL, "{0}", sLatest)
    UpdateChecker_GUI.Show vbModal
    Unload UpdateChecker_GUI
End Sub

' Downloads ARES.mvba to %TEMP%, launches an elevated PowerShell script that waits for
' MicroStation to close then copies the file (UAC prompt expected), then quits.
' Called by UpdateChecker_GUI.cmdYes_Click.
Public Sub DownloadAndInstall()
    On Error GoTo ErrorHandler

    Dim oHttp As Object
    Dim oStream As Object
    Dim oShell As Object
    Dim sTempMvba As String
    Dim sTempPs As String
    Dim iFile As Integer

    ShowStatus GetTranslation("UpdateDownloading")

    ' [1] Download ARES.mvba to %TEMP% — avoids both the folder permission and the file lock
    sTempMvba = Environ("TEMP") & "\ARES_update.mvba"
    sTempPs = Environ("TEMP") & "\ARES_update.ps1"

    Set oHttp = CreateObject("MSXML2.XMLHTTP")
    oHttp.Open "GET", gsUpdateDownloadUrl, False
    oHttp.setRequestHeader "User-Agent", "ARES-MVBA"
    oHttp.send

    If oHttp.Status <> 200 Then GoTo ErrorHandler

    Set oStream = CreateObject("ADODB.Stream")
    oStream.Type = 1 ' Binary
    oStream.Open
    oStream.Write oHttp.responseBody
    oStream.SaveToFile sTempMvba, 2 ' adSaveCreateOverWrite
    oStream.Close
    Set oStream = Nothing
    Set oHttp = Nothing

    ' [2] Write a PowerShell script that retries until the lock is released then copies
    iFile = FreeFile
    Open sTempPs For Output As #iFile
    Print #iFile, "do {"
    Print #iFile, "    Start-Sleep -Seconds 2"
    Print #iFile, "    try {"
    Print #iFile, "        Copy-Item -Path '" & sTempMvba & "' -Destination '" & ARES_MVBA_PATH & "' -Force -ErrorAction Stop"
    Print #iFile, "        $done = $true"
    Print #iFile, "    } catch { $done = $false }"
    Print #iFile, "} while (-not $done)"
    Print #iFile, "Remove-Item -Path '" & sTempMvba & "' -Force -ErrorAction SilentlyContinue"
    Print #iFile, "Remove-Item -Path '" & sTempPs & "' -Force -ErrorAction SilentlyContinue"
    Close #iFile

    ' [3] Launch the script elevated (triggers UAC) — hidden window, no wait
    Set oShell = CreateObject("Shell.Application")
    oShell.ShellExecute "powershell.exe", _
        "-ExecutionPolicy Bypass -WindowStyle Hidden -File """ & sTempPs & """", _
        "", "runas", 0
    Set oShell = Nothing

    ' [4] Quit MicroStation — script will copy once the lock is released
    Application.Quit
    Exit Sub

ErrorHandler:
    Set oStream = Nothing
    Set oHttp = Nothing
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "UpdateChecker.DownloadAndInstall"
    MsgBox GetTranslation("UpdateDownloadFailed"), vbCritical, "ARES"
End Sub
