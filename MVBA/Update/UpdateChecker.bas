' Module: UpdateChecker
' Description: Checks for available ARES updates via the GitHub Releases API.
'              The installed version is written to the Windows Registry by the installer.
'              User notification preferences are stored as MicroStation configuration variables.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (via BootLoader)
Option Explicit

Private Const ARES_REGISTRY_KEY As String = "HKCU\Software\ARES\Version"
Private Const ARES_GITHUB_API_URL As String = "https://api.github.com/repos/Asketyll/ARES/releases/latest"
Private Const ARES_GITHUB_DOWNLOAD_URL As String = "https://github.com/Asketyll/ARES/releases/download/v.{0}/ARES.mvba"
Private Const ARES_MVBA_PATH As String = "C:\ARES\ARES.mvba"
Private msLatestVersion As String
Private msDownloadUrl As String

' Read-only accessors for UpdateChecker_GUI
Public Function GetUpdateLatestVersion() As String
    GetUpdateLatestVersion = msLatestVersion
End Function

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

' Returns the latest release tag from the GitHub API (without leading "v.").
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

    ' Strip leading "v." if present (e.g. "v.1.2.3" -> "1.2.3")
    If Left(sTagName, 2) = "v." Then sTagName = Mid(sTagName, 3)

    GetLatestVersionFromGitHub = sTagName
    Exit Function

ErrorHandler:
    Set oHttp = Nothing
    GetLatestVersionFromGitHub = ""
End Function

' Compares two "major.minor.patch" version strings.
' Returns  1 if v1 > v2, -1 if v1 < v2, 0 if equal.
' Extra segments are treated as 0 (e.g. "1.2" = "1.2.0").
Private Function CompareVersions(ByVal v1 As String, ByVal v2 As String) As Integer
    On Error GoTo ErrorHandler

    Dim a1() As String
    Dim a2() As String
    Dim i As Integer
    Dim n As Integer
    Dim p1 As Long
    Dim p2 As Long

    a1 = Split(v1, ".")
    a2 = Split(v2, ".")
    n = IIf(UBound(a1) > UBound(a2), UBound(a1), UBound(a2))

    For i = 0 To n
        p1 = IIf(i <= UBound(a1), CLng(a1(i)), 0)
        p2 = IIf(i <= UBound(a2), CLng(a2(i)), 0)
        If p1 > p2 Then CompareVersions = 1 : Exit Function
        If p1 < p2 Then CompareVersions = -1 : Exit Function
    Next i

    CompareVersions = 0
    Exit Function

ErrorHandler:
    ' Malformed version string — treat as equal (no update proposed)
    CompareVersions = 0
End Function

' Returns True only if the version string contains digits and dots exclusively.
' Whitelist validation — rejects any character that could inject PowerShell commands.
Private Function IsValidVersion(ByVal sVersion As String) As Boolean
    Dim i As Integer
    Dim c As String

    If Len(sVersion) = 0 Then Exit Function

    For i = 1 To Len(sVersion)
        c = Mid(sVersion, i, 1)
        If c <> "." And (c < "0" Or c > "9") Then Exit Function
    Next i

    IsValidVersion = True
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

    ' [3b] Reject any version string that is not strictly "digits and dots" — PS injection guard
    If Not IsValidVersion(sLatest) Then Exit Sub

    ' [4] Already up to date, or latest is older (no downgrade)
    If CompareVersions(sLatest, sInstalled) <= 0 Then Exit Sub

    ' [5] User previously chose to ignore this specific version
    If ARESConfig.ARES_UPDATE_IGNORE_VERSION.Value = sLatest Then Exit Sub

    ' [6] Show notification dialog
    ShowUpdateDialog sInstalled, sLatest
End Sub

' Shows the update notification dialog (UpdateChecker_GUI).
' Buttons: Update | Ignore this version | Ignore all updates
Private Sub ShowUpdateDialog(ByVal sInstalled As String, ByVal sLatest As String)
    On Error Resume Next

    msLatestVersion = sLatest
    msDownloadUrl = Replace(ARES_GITHUB_DOWNLOAD_URL, "{0}", sLatest)
    UpdateChecker_GUI.Show vbModal
    Unload UpdateChecker_GUI
End Sub

' Downloads ARES.mvba via an elevated PowerShell script that:
'   - Runs from an isolated temp folder
'   - Retries the copy up to 30 times, ~1 min
'   - Updates the registry version on success
'   - Cleans up the temp folder on exit
' Called by UpdateChecker_GUI.cmdYes_Click.
Public Sub DownloadAndInstall()
    On Error GoTo ErrorHandler

    Dim oFso As Object
    Dim oShell As Object
    Dim sTempDir As String
    Dim sTempMvba As String
    Dim sTempPs As String
    Dim iFile As Integer

    ShowStatus GetTranslation("UpdateDownloading")
    Set oFso = CreateObject("Scripting.FileSystemObject")
    sTempDir = Environ("TEMP") & "\" & oFso.GetTempName()
    oFso.CreateFolder sTempDir
    Set oFso = Nothing

    sTempMvba = sTempDir & "\ARES_update.mvba"
    sTempPs = sTempDir & "\ARES_update.ps1"

    ' [2] Write the PowerShell script
    iFile = FreeFile
    Open sTempPs For Output As #iFile
    Print #iFile, "$src      = '" & sTempMvba & "'"
    Print #iFile, "$dst      = '" & ARES_MVBA_PATH & "'"
    Print #iFile, "$url      = '" & msDownloadUrl & "'"
    Print #iFile, "$dir      = '" & sTempDir & "'"
    Print #iFile, "$version  = '" & msLatestVersion & "'"
    Print #iFile, "$maxRetry = 30"
    Print #iFile, "$attempt  = 0"
    Print #iFile, "Invoke-WebRequest -Uri $url -OutFile $src -UseBasicParsing"
    Print #iFile, "# Verify SHA256 integrity before installing"
    Print #iFile, "try {"
    Print #iFile, "    $hashUrl  = ""$url.sha256"""
    Print #iFile, "    $hashFile = ""$src.sha256"""
    Print #iFile, "    Invoke-WebRequest -Uri $hashUrl -OutFile $hashFile -UseBasicParsing"
    Print #iFile, "    $expectedHash = ((Get-Content $hashFile -Raw).Trim() -split '\s+')[0].ToUpper()"
    Print #iFile, "    $actualHash   = (Get-FileHash -Path $src -Algorithm SHA256).Hash.ToUpper()"
    Print #iFile, "    if ($expectedHash -ne $actualHash) {"
    Print #iFile, "        Add-Type -AssemblyName System.Windows.Forms"
    Print #iFile, "        [System.Windows.Forms.MessageBox]::Show('ARES : hash verification failed. Update aborted.`nARES : v" & Chr(233) & "rification du fichier " & Chr(233) & "chou" & Chr(233) & "e. Mise " & Chr(224) & " jour annul" & Chr(233) & "e.', 'ARES', 0, 48) | Out-Null"
    Print #iFile, "        Remove-Item -Path $dir -Recurse -Force -ErrorAction SilentlyContinue"
    Print #iFile, "        exit 1"
    Print #iFile, "    }"
    Print #iFile, "} catch {"
    Print #iFile, "    Add-Type -AssemblyName System.Windows.Forms"
    Print #iFile, "    [System.Windows.Forms.MessageBox]::Show('ARES : could not verify file integrity. Update aborted.`nARES : impossible de v" & Chr(233) & "rifier l''int" & Chr(233) & "grit" & Chr(233) & " du fichier. Mise " & Chr(224) & " jour annul" & Chr(233) & "e.', 'ARES', 0, 48) | Out-Null"
    Print #iFile, "    Remove-Item -Path $dir -Recurse -Force -ErrorAction SilentlyContinue"
    Print #iFile, "    exit 1"
    Print #iFile, "}"
    Print #iFile, "do {"
    Print #iFile, "    Start-Sleep -Seconds 2"
    Print #iFile, "    $attempt++"
    Print #iFile, "    try {"
    Print #iFile, "        Copy-Item -Path $src -Destination $dst -Force -ErrorAction Stop"
    Print #iFile, "        if (-not (Test-Path 'HKCU:\Software\ARES')) { New-Item -Path 'HKCU:\Software\ARES' -Force | Out-Null }"
    Print #iFile, "        Set-ItemProperty -Path 'HKCU:\Software\ARES' -Name 'Version' -Value $version"
    Print #iFile, "        $done = $true"
    Print #iFile, "    } catch { $done = $false }"
    Print #iFile, "} while (-not $done -and $attempt -lt $maxRetry)"
    Print #iFile, "Remove-Item -Path $dir -Recurse -Force -ErrorAction SilentlyContinue"
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
    ' Clean up the isolated temp folder if it was created
    On Error Resume Next
    If Len(sTempDir) > 0 Then
        Dim oFsoClean As Object
        Set oFsoClean = CreateObject("Scripting.FileSystemObject")
        If oFsoClean.FolderExists(sTempDir) Then oFsoClean.DeleteFolder sTempDir, True
        Set oFsoClean = Nothing
    End If
    On Error GoTo 0
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "UpdateChecker.DownloadAndInstall"
    MsgBox GetTranslation("UpdateDownloadFailed"), vbCritical, "ARES"
End Sub
