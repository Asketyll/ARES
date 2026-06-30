' Module: UpdateChecker
' Description: Checks for available ARES updates via the GitHub Releases API and installs every
'              payload asset of the latest release (ARES.mvba + resource files such as the
'              custom-property .dgnlib). The installed version is written to the Windows Registry.
'              User notification preferences are stored as MicroStation configuration variables.
'
'              Placement rule (mirrors the installer's CopyResources): ".mvba" goes to C:\ARES,
'              every other asset goes to C:\ARES\Rsc. Each asset is SHA-256 verified (digest from
'              the GitHub API) before anything is copied; the copy waits for MicroStation to release
'              its file locks (retry loop) after the application quits.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants, ARESConfigClass (via BootLoader)
Option Explicit

Private Const ARES_REGISTRY_KEY As String = "HKCU\Software\ARES\Version"
Private Const ARES_GITHUB_API_URL As String = "https://api.github.com/repos/Asketyll/ARES/releases/latest"
Private Const ARES_DOWNLOAD_BASE As String = "https://github.com/Asketyll/ARES/releases/download/v."  ' + <version>/<assetName>
Private Const ARES_INSTALL_DIR As String = "C:\ARES"
Private Const ARES_RSC_DIR As String = "C:\ARES\Rsc"
Private Const MAX_ASSETS As Long = 31

' One downloadable release asset (name + SHA-256 hex digest). URL is derived from the version + name.
Private Type AssetInfo
    Name As String
    Hash As String
End Type

Private msLatestVersion As String
Private mAssets(0 To MAX_ASSETS) As AssetInfo
Private mnAssetCount As Long

' Read-only accessor for UpdateChecker_GUI
Public Function GetUpdateLatestVersion() As String
    GetUpdateLatestVersion = msLatestVersion
End Function

' Returns the installed version from the Windows Registry (written by the installer).
' Returns "" if the key is absent or unreadable - callers must handle this and skip the check.
Public Function GetInstalledVersion() As String
    On Error Resume Next

    Dim oShell As Object
    Dim sVersion As String

    Set oShell = CreateObject("WScript.Shell")
    sVersion = oShell.RegRead(ARES_REGISTRY_KEY)
    Set oShell = Nothing

    If Err.Number <> 0 Then Err.Clear : Exit Function
    GetInstalledVersion = Trim(sVersion)
End Function

' Returns the latest release tag from the GitHub API (without leading "v.") and parses the release
' assets (name + digest) into mAssets. Returns "" on network/HTTP/parse failure - never raises.
Public Function GetLatestVersionFromGitHub() As String
    On Error GoTo ErrorHandler

    Dim oHttp As Object
    Dim sResponse As String
    Dim sTagName As String
    Dim nStart As Long
    Dim nEnd As Long

    mnAssetCount = 0

    Set oHttp = CreateObject("MSXML2.XMLHTTP")
    oHttp.Open "GET", ARES_GITHUB_API_URL, False
    oHttp.setRequestHeader "User-Agent", "ARES-MVBA"
    oHttp.send

    If oHttp.Status <> 200 Then GoTo ErrorHandler

    sResponse = oHttp.responseText
    Set oHttp = Nothing

    ' Extract tag_name value from JSON response
    nStart = InStr(sResponse, """tag_name""")
    If nStart = 0 Then GoTo ErrorHandler

    nStart = InStr(nStart, sResponse, ":")
    If nStart = 0 Then GoTo ErrorHandler

    nStart = InStr(nStart, sResponse, """")
    If nStart = 0 Then GoTo ErrorHandler
    nStart = nStart + 1

    nEnd = InStr(nStart, sResponse, """")
    If nEnd = 0 Then GoTo ErrorHandler

    sTagName = Mid(sResponse, nStart, nEnd - nStart)

    ' Strip leading "v." if present (e.g. "v.1.2.3" -> "1.2.3")
    If Left(sTagName, 2) = "v." Then sTagName = Mid(sTagName, 3)

    ParseAssets sResponse

    GetLatestVersionFromGitHub = sTagName
    Exit Function

ErrorHandler:
    Set oHttp = Nothing
    mnAssetCount = 0
    GetLatestVersionFromGitHub = ""
End Function

' Parse the "assets" array of the API response into mAssets (name + 64-hex sha256). Only assets
' with a safe filename AND a valid digest are kept (anything unverifiable is ignored). Scanning is
' bounded to the assets section so the release name / body text cannot contaminate the parse.
Private Sub ParseAssets(ByVal sResponse As String)
    On Error GoTo ErrorHandler

    Dim nAssetsStart As Long, nAssetsEnd As Long
    Dim nNamePos As Long, nNextName As Long, nDigestPos As Long, nShaPos As Long, nValEnd As Long
    Dim sName As String, sHash As String

    mnAssetCount = 0

    nAssetsStart = InStr(sResponse, """assets""")
    If nAssetsStart = 0 Then Exit Sub

    ' assets is followed by "tarball_url" / "zipball_url" in the release JSON.
    nAssetsEnd = InStr(nAssetsStart, sResponse, """tarball_url""")
    If nAssetsEnd = 0 Then nAssetsEnd = InStr(nAssetsStart, sResponse, """zipball_url""")
    If nAssetsEnd = 0 Then nAssetsEnd = Len(sResponse)

    nNamePos = nAssetsStart
    Do
        nNamePos = InStr(nNamePos + 1, sResponse, """name""")
        If nNamePos = 0 Or nNamePos >= nAssetsEnd Then Exit Do

        sName = ExtractQuotedValue(sResponse, nNamePos, nAssetsEnd)

        nNextName = InStr(nNamePos + 1, sResponse, """name""")
        If nNextName = 0 Or nNextName > nAssetsEnd Then nNextName = nAssetsEnd

        ' digest belonging to THIS asset (between its name and the next asset's name)
        sHash = ""
        nDigestPos = InStr(nNamePos, sResponse, """digest""")
        If nDigestPos > 0 And nDigestPos < nNextName Then
            nShaPos = InStr(nDigestPos, sResponse, "sha256:")
            If nShaPos > 0 And nShaPos < nNextName Then
                nValEnd = InStr(nShaPos, sResponse, """")
                If nValEnd > 0 Then sHash = Mid(sResponse, nShaPos + 7, nValEnd - (nShaPos + 7))
            End If
        End If
        sHash = LCase(Trim(sHash))

        If IsSafeAssetName(sName) And Len(sHash) = 64 And IsHex(sHash) Then
            If mnAssetCount <= UBound(mAssets) Then
                mAssets(mnAssetCount).Name = sName
                mAssets(mnAssetCount).Hash = sHash
                mnAssetCount = mnAssetCount + 1
            End If
        End If

        nNamePos = nNextName
    Loop While nNamePos < nAssetsEnd

    Exit Sub

ErrorHandler:
    mnAssetCount = 0
End Sub

' Extract the quoted string value that follows a "key" token (bounded by limit). Values handled
' here (asset names, hashes) never contain quotes, so simple delimiter scanning is sufficient.
Private Function ExtractQuotedValue(ByVal s As String, ByVal keyPos As Long, ByVal limit As Long) As String
    Dim nColon As Long, nOpen As Long, nClose As Long

    nColon = InStr(keyPos, s, ":")
    If nColon = 0 Or nColon > limit Then Exit Function

    nOpen = InStr(nColon, s, """")
    If nOpen = 0 Or nOpen > limit Then Exit Function
    nOpen = nOpen + 1

    nClose = InStr(nOpen, s, """")
    If nClose = 0 Or nClose > limit Then Exit Function

    ExtractQuotedValue = Mid(s, nOpen, nClose - nOpen)
End Function

' True if the latest release exposes an installable, verified ARES.mvba (the minimum to offer an update).
Private Function HasInstallableMvba() As Boolean
    Dim i As Long
    For i = 0 To mnAssetCount - 1
        If LCase(mAssets(i).Name) = "ares.mvba" Then
            HasInstallableMvba = True
            Exit Function
        End If
    Next i
End Function

' Whitelist a downloadable asset name: non-empty, <=128 chars, only [A-Za-z0-9._-], no "..".
' Guards both path traversal in the copy target and PowerShell string injection.
Private Function IsSafeAssetName(ByVal s As String) As Boolean
    Dim i As Long, ch As String

    If Len(s) = 0 Or Len(s) > 128 Then Exit Function
    If InStr(s, "..") > 0 Then Exit Function

    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If Not ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or _
                (ch >= "0" And ch <= "9") Or ch = "." Or ch = "_" Or ch = "-") Then Exit Function
    Next i

    IsSafeAssetName = True
End Function

' True if the string is non-empty and contains only hexadecimal characters.
Private Function IsHex(ByVal s As String) As Boolean
    Dim i As Long, ch As String

    If Len(s) = 0 Then Exit Function

    For i = 1 To Len(s)
        ch = LCase(Mid(s, i, 1))
        If Not ((ch >= "0" And ch <= "9") Or (ch >= "a" And ch <= "f")) Then Exit Function
    Next i

    IsHex = True
End Function

' Compares two "major.minor.patch" version strings. Returns 1 if v1 > v2, -1 if v1 < v2, 0 if equal.
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
    CompareVersions = 0
End Function

' True only if the version string contains digits and dots exclusively (PS injection guard).
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

' Manually checks for an available update and shows the dialog unconditionally
' (bypasses mute and ignore-version preferences).
Public Sub CheckForUpdateManual()
    On Error Resume Next

    Dim sInstalled As String
    Dim sLatest As String

    sInstalled = GetInstalledVersion()
    If Len(sInstalled) = 0 Then ShowStatus GetTranslation("UpdateCheckFailed") : Exit Sub

    sLatest = GetLatestVersionFromGitHub()
    If Len(sLatest) = 0 Then ShowStatus GetTranslation("UpdateCheckFailed") : Exit Sub
    If Not IsValidVersion(sLatest) Then ShowStatus GetTranslation("UpdateCheckFailed") : Exit Sub
    If Not HasInstallableMvba() Then ShowStatus GetTranslation("UpdateCheckFailed") : Exit Sub

    If CompareVersions(sLatest, sInstalled) <= 0 Then
        ShowStatus GetTranslation("UpdateAlreadyUpToDate")
        Exit Sub
    End If

    ShowUpdateDialog sInstalled, sLatest
End Sub

' Checks for an available update and notifies the user if one is found.
' Silently exits if the network is unavailable, the check fails, or preferences suppress it.
Public Sub CheckForUpdate()
    On Error Resume Next

    Dim sInstalled As String
    Dim sLatest As String

    ' [1] Permanently muted?
    If ARESConfig.ARES_UPDATE_MUTE.Value = "True" Then Exit Sub

    ' [2] Installed version from registry - empty means not installed via installer, skip
    sInstalled = GetInstalledVersion()
    If Len(sInstalled) = 0 Then Exit Sub

    ' [3] Latest version from GitHub - empty means network unavailable
    sLatest = GetLatestVersionFromGitHub()
    If Len(sLatest) = 0 Then Exit Sub

    ' [3b] PS injection guard - reject non "digits and dots"
    If Not IsValidVersion(sLatest) Then Exit Sub

    ' [3c] Require a verifiable ARES.mvba in the release - abort silently otherwise
    If Not HasInstallableMvba() Then Exit Sub

    ' [4] Already up to date, or latest is older (no downgrade)
    If CompareVersions(sLatest, sInstalled) <= 0 Then Exit Sub

    ' [5] User previously chose to ignore this specific version
    If ARESConfig.ARES_UPDATE_IGNORE_VERSION.Value = sLatest Then Exit Sub

    ' [6] Notify
    ShowUpdateDialog sInstalled, sLatest
End Sub

' Shows the update notification dialog (UpdateChecker_GUI).
Private Sub ShowUpdateDialog(ByVal sInstalled As String, ByVal sLatest As String)
    On Error Resume Next

    msLatestVersion = sLatest
    UpdateChecker_GUI.Show vbModal
    Unload UpdateChecker_GUI
End Sub

' Downloads and installs every payload asset of the latest release via an elevated PowerShell script:
'   - downloads each asset to an isolated temp folder and SHA-256 verifies it (aborts all on mismatch)
'   - ".mvba" -> C:\ARES, every other asset -> C:\ARES\Rsc
'   - retries the copy of the whole set up to 30 times (~1 min) while MicroStation releases its locks
'   - updates the registry version on success, then cleans up
' Called by UpdateChecker_GUI.cmdYes_Click.
Public Sub DownloadAndInstall()
    On Error GoTo ErrorHandler

    Dim oFso As Object
    Dim oShell As Object
    Dim sTempDir As String
    Dim sTempPs As String
    Dim iFile As Integer
    Dim i As Long
    Dim sName As String
    Dim sUrl As String
    Dim sTarget As String

    ' Defensive re-validation (these are already checked before the dialog is shown).
    If Not IsValidVersion(msLatestVersion) Then Exit Sub
    If mnAssetCount = 0 Then Exit Sub

    ShowStatus GetTranslation("UpdateDownloading")
    Set oFso = CreateObject("Scripting.FileSystemObject")
    sTempDir = Environ("TEMP") & "\" & oFso.GetTempName()
    oFso.CreateFolder sTempDir
    Set oFso = Nothing

    sTempPs = sTempDir & "\ARES_update.ps1"

    iFile = FreeFile
    Open sTempPs For Output As #iFile

    Print #iFile, "$dir     = '" & sTempDir & "'"
    Print #iFile, "$version = '" & msLatestVersion & "'"
    Print #iFile, "$rsc     = '" & ARES_RSC_DIR & "'"
    Print #iFile, "$assets = @("

    For i = 0 To mnAssetCount - 1
        sName = mAssets(i).Name
        sUrl = ARES_DOWNLOAD_BASE & msLatestVersion & "/" & sName
        If LCase(Right(sName, 5)) = ".mvba" Then
            sTarget = ARES_INSTALL_DIR & "\" & sName
        Else
            sTarget = ARES_RSC_DIR & "\" & sName
        End If
        Print #iFile, "  @{ Name = '" & sName & "'; Url = '" & sUrl & "'; Hash = '" & mAssets(i).Hash & "'; Target = '" & sTarget & "' }" & IIf(i < mnAssetCount - 1, ",", "")
    Next i

    Print #iFile, ")"

    ' [1] Download + verify every asset; abort all if any hash mismatches.
    Print #iFile, "foreach ($a in $assets) {"
    Print #iFile, "    $tmp = Join-Path $dir $a.Name"
    Print #iFile, "    Invoke-WebRequest -Uri $a.Url -OutFile $tmp -UseBasicParsing"
    Print #iFile, "    $actual = (Get-FileHash -Path $tmp -Algorithm SHA256).Hash.ToLower()"
    Print #iFile, "    if ($a.Hash -ne $actual) {"
    Print #iFile, "        Add-Type -AssemblyName System.Windows.Forms"
    Print #iFile, "        [System.Windows.Forms.MessageBox]::Show('ARES : hash verification failed. Update aborted.`nARES : v" & Chr(233) & "rification du fichier " & Chr(233) & "chou" & Chr(233) & "e. Mise " & Chr(224) & " jour annul" & Chr(233) & "e.', 'ARES', 0, 48) | Out-Null"
    Print #iFile, "        Remove-Item -Path $dir -Recurse -Force -ErrorAction SilentlyContinue"
    Print #iFile, "        exit 1"
    Print #iFile, "    }"
    Print #iFile, "}"

    ' [2] Ensure the Rsc folder exists.
    Print #iFile, "if (-not (Test-Path $rsc)) { New-Item -ItemType Directory -Force -Path $rsc | Out-Null }"

    ' [3] Copy the whole set, retrying while MicroStation still holds the file locks.
    Print #iFile, "$attempt = 0"
    Print #iFile, "$done    = $false"
    Print #iFile, "do {"
    Print #iFile, "    Start-Sleep -Seconds 2"
    Print #iFile, "    $attempt++"
    Print #iFile, "    try {"
    Print #iFile, "        foreach ($a in $assets) { Copy-Item -Path (Join-Path $dir $a.Name) -Destination $a.Target -Force -ErrorAction Stop }"
    Print #iFile, "        $done = $true"
    Print #iFile, "    } catch { $done = $false }"
    Print #iFile, "} while (-not $done -and $attempt -lt 30)"

    ' [4] Record the new version on success.
    Print #iFile, "if ($done) {"
    Print #iFile, "    if (-not (Test-Path 'HKCU:\Software\ARES')) { New-Item -Path 'HKCU:\Software\ARES' -Force | Out-Null }"
    Print #iFile, "    Set-ItemProperty -Path 'HKCU:\Software\ARES' -Name 'Version' -Value $version"
    Print #iFile, "}"
    Print #iFile, "Remove-Item -Path $dir -Recurse -Force -ErrorAction SilentlyContinue"

    Close #iFile

    ' Launch elevated (triggers UAC) - hidden window, no wait.
    Set oShell = CreateObject("Shell.Application")
    oShell.ShellExecute "powershell.exe", _
        "-ExecutionPolicy Bypass -WindowStyle Hidden -File """ & sTempPs & """", _
        "", "runas", 0
    Set oShell = Nothing

    ' Quit MicroStation - the script copies once the locks are released.
    Application.Quit
    Exit Sub

ErrorHandler:
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
