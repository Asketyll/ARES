' Module: ConfigThemes
' Description: Configuration themes - named business configurations kept as .cfg files in the theme folder:
'              list them, check a name, load one, save the current one, and make a loaded configuration
'              govern ARES at once.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, ARESConstants, LangManager, ErrorHandlerClass, PropertyTagging, PropertyCalculation,
'               PropertyActuator, Command
Option Explicit

Private Const ARES_ROOT_FOLDER As String = "C:\ARES"
Private Const THEME_EXTENSION As String = ".cfg"
Private Const MAX_THEME_PATH As Long = 259

' C:\ARES\Theme with an e-grave (ChrW 232): no non-ASCII literal, so the path never depends on how the source is
' encoded or imported.
Public Function ThemeFolder() As String
    On Error GoTo ErrorHandler
    ThemeFolder = ARES_ROOT_FOLDER & "\Th" & ChrW(232) & "me"
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes.ThemeFolder"
End Function

' Create the theme folder, and C:\ARES before it, when absent. A failure is logged and returns False.
Public Function EnsureThemeFolder() As Boolean
    On Error GoTo ErrorHandler

    EnsureThemeFolder = False
    If Not FolderExists(ARES_ROOT_FOLDER) Then MkDir ARES_ROOT_FOLDER
    If Not FolderExists(ThemeFolder()) Then MkDir ThemeFolder()
    EnsureThemeFolder = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes.EnsureThemeFolder"
    EnsureThemeFolder = False
End Function

' The theme names (file names without .cfg), in case-insensitive order; returns their count. An absent or empty
' folder gives 0, and is not a fault.
Public Function ListThemes(ByRef sNames() As String) As Long
    On Error GoTo ErrorHandler

    Dim sFile As String
    Dim sHeld As String
    Dim nFound As Long
    Dim i As Long
    Dim j As Long

    ListThemes = 0
    ReDim sNames(0 To 0)
    If Not FolderExists(ThemeFolder()) Then Exit Function

    ' Every Dir result is collected before anything else runs: any other Dir call restarts the enumeration.
    ' "*.cfg" would also match a longer extension through its 8.3 short name, hence the filter on the name.
    sFile = Dir(ThemeFolder() & "\*", vbNormal)
    Do While Len(sFile) > 0
        If Len(sFile) > Len(THEME_EXTENSION) Then
            If LCase(Right(sFile, Len(THEME_EXTENSION))) = THEME_EXTENSION Then
                ReDim Preserve sNames(0 To nFound)
                sNames(nFound) = Left(sFile, Len(sFile) - Len(THEME_EXTENSION))
                nFound = nFound + 1
            End If
        End If
        sFile = Dir()
    Loop

    ' Insertion sort: a theme folder holds a handful of files.
    For i = 1 To nFound - 1
        sHeld = sNames(i)
        j = i - 1
        Do While j >= 0
            If StrComp(sNames(j), sHeld, vbTextCompare) <= 0 Then Exit Do
            sNames(j + 1) = sNames(j)
            j = j - 1
        Loop
        sNames(j + 1) = sHeld
    Next i

    ListThemes = nFound
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes.ListThemes"
    ReDim sNames(0 To 0)
    ListThemes = 0
End Function

' A theme name Windows can store as <name>.cfg in the theme folder: not blank, none of \ / : * ? " < > | or a
' control character, no final dot, not a device name (CON, PRN, AUX, NUL, COM0-9, LPT0-9, and COM/LPT followed by
' a superscript 1, 2 or 3), every character held by the ANSI code page (the file APIs used here are ANSI), and the
' full path within 259 characters.
Public Function IsValidThemeName(ByVal sName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim i As Long
    Dim nCode As Long
    Dim sBase As String

    IsValidThemeName = False
    If Len(Trim(sName)) = 0 Then Exit Function
    If Right(sName, 1) = "." Then Exit Function

    For i = 1 To Len(sName)
        nCode = AscW(Mid(sName, i, 1))
        If nCode >= 0 And nCode < 32 Then Exit Function
        If InStr("\/:*?""<>|", Mid(sName, i, 1)) > 0 Then Exit Function
    Next i

    ' Windows reserves a device name whatever extension follows it.
    sBase = UCase(RTrim(Split(sName, ".")(0)))
    Select Case sBase
        Case "CON", "PRN", "AUX", "NUL", _
             "COM0", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", _
             "LPT0", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9", _
             "COM" & ChrW(185), "COM" & ChrW(178), "COM" & ChrW(179), _
             "LPT" & ChrW(185), "LPT" & ChrW(178), "LPT" & ChrW(179)
            Exit Function
    End Select

    If StrConv(StrConv(sName, vbFromUnicode), vbUnicode) <> sName Then Exit Function
    If Len(ThemePath(sName)) > MAX_THEME_PATH Then Exit Function

    IsValidThemeName = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes.IsValidThemeName"
    IsValidThemeName = False
End Function

' Load a theme: its business variables replace the current ones (station lines in the file are ignored, absent
' variables keep their value), then they govern ARES at once. Every outcome goes to the status bar; a file that
' cannot be read, or holds no business line, changes nothing - the displayed theme included.
Public Function LoadTheme(ByVal sName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim bReadable As Boolean
    Dim nUnknown As Long
    Dim sFileVersion As String
    Dim sWarnings As String

    LoadTheme = False
    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    If Not ARESConfig.LoadConfig(ThemePath(sName), True, True, bReadable, nUnknown, sFileVersion) Then
        If bReadable Then
            LangManager.ShowStatusText GetTranslation("ConfigThemeNoSetting", sName)
        Else
            LangManager.ShowStatusText GetTranslation("ConfigThemeUnreadable", sName)
        End If
        Exit Function
    End If

    ARESConfig.ARES_THEME_CURRENT.Value = sName
    ApplyLoadedConfiguration

    sWarnings = DescribeLoadWarnings(nUnknown, sFileVersion)
    If Len(sWarnings) > 0 Then
        LangManager.ShowStatusText GetTranslation("ConfigThemeLoadedWarnings", sName, sWarnings)
    Else
        LangManager.ShowStatusText GetTranslation("ConfigThemeLoaded", sName)
    End If
    LoadTheme = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes.LoadTheme"
    LoadTheme = False
End Function

' Save every business variable as the theme sName, which becomes the displayed theme. An existing theme is
' overwritten only on Yes. Refusals and the outcome go to the status bar; a write fault is also logged.
Public Function SaveTheme(ByVal sName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim sPath As String
    Dim sOnDisk As String

    SaveTheme = False
    If Not LangManager.IsInit Then LangManager.InitializeTranslations

    sName = Trim(sName)
    If Not IsValidThemeName(sName) Then
        LangManager.ShowStatusT "ConfigThemeNameInvalid"
        Exit Function
    End If

    If Not EnsureThemeFolder() Then
        LangManager.ShowStatusText GetTranslation("ConfigThemeSaveFailed", sName)
        Exit Function
    End If

    sPath = ThemePath(sName)
    If FileExists(sPath) Then
        ' File names are case-insensitive: the file keeps its own spelling, so the theme is named after it.
        sOnDisk = Dir(sPath)
        If Len(sOnDisk) > Len(THEME_EXTENSION) Then sName = Left(sOnDisk, Len(sOnDisk) - Len(THEME_EXTENSION))
        If MsgBox(GetTranslation("ConfigThemeOverwritePrompt", sName), vbYesNo + vbQuestion, "ARES") <> vbYes Then
            LangManager.ShowStatusT "ConfigOperationCancelled"
            Exit Function
        End If
    End If

    If Not ARESConfig.ExportConfig(sPath, True) Then
        LangManager.ShowStatusText GetTranslation("ConfigThemeSaveFailed", sName)
        Exit Function
    End If

    ARESConfig.ARES_THEME_CURRENT.Value = sName
    LangManager.ShowStatusText GetTranslation("ConfigThemeSaved", sName)
    SaveTheme = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes.SaveTheme"
    LangManager.ShowStatusText GetTranslation("ConfigThemeSaveFailed", sName)
    SaveTheme = False
End Function

' Make a configuration just loaded (a theme or ImportARESConfig) govern ARES without a restart: drop the parsed
' rule caches, and show the loaded values in every open ARES form.
Public Sub ApplyLoadedConfiguration()
    On Error GoTo ErrorHandler
    PropertyTagging.RefreshRules
    PropertyCalculation.RefreshCalcRules
    PropertyActuator.RefreshActuatorState
    Command.RefreshOpenForms
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes.ApplyLoadedConfiguration"
End Sub

' The warnings of a load as one translated clause, "" when there is none: unknown lines ignored, and a
' "# Version:" header other than ARES_CONFIG_VERSION.
Public Function DescribeLoadWarnings(ByVal nUnknown As Long, ByVal sFileVersion As String) As String
    On Error GoTo ErrorHandler

    Dim sText As String

    If nUnknown > 0 Then sText = GetTranslation("ConfigWarnUnknownLines", nUnknown)
    If Len(sFileVersion) > 0 Then
        If sFileVersion <> ARES_CONFIG_VERSION Then
            If Len(sText) > 0 Then sText = sText & "; "
            sText = sText & GetTranslation("ConfigWarnVersion", sFileVersion, ARES_CONFIG_VERSION)
        End If
    End If
    DescribeLoadWarnings = sText
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ConfigThemes.DescribeLoadWarnings"
    DescribeLoadWarnings = ""
End Function

Private Function ThemePath(ByVal sName As String) As String
    ThemePath = ThemeFolder() & "\" & sName & THEME_EXTENSION
End Function

' GetAttr under a local trap: an absent path is an expected answer, not a fault.
Private Function FolderExists(ByVal sPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(sPath) And vbDirectory) = vbDirectory)
    If Err.Number <> 0 Then FolderExists = False
    Err.Clear
End Function

Private Function FileExists(ByVal sPath As String) As Boolean
    On Error Resume Next
    FileExists = ((GetAttr(sPath) And vbDirectory) = 0)
    If Err.Number <> 0 Then FileExists = False
    Err.Clear
End Function
