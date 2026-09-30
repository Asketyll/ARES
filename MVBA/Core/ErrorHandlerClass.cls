' Class Module: ErrorHandlerClass
' Description: Centralized error handling and logging for ARES application
' License: This project is licensed under the AGPL-3.0.
' Dependencies: CallStackClass (read-only, via the global BootLoader.CallStack - see
'               FormatErrorMessage; CallStackClass itself has no dependency back on this class),
'               Messaging (Message Center), LangManager (read only when already initialized)
Option Explicit

' === CONSTANTS ===
Private Const MAX_LOG_FILE_SIZE As Long = 1048576    ' 1 MB
' English fallback of the LogFileCreated translation, used while LangManager is not initialized
Private Const LOG_FILE_CREATED_EN As String = "ARES created a log file: {0}"

' === PRIVATE MEMBERS ===
Private msLogFilePath As String
' Set True when a real fault (Number > 0) is logged. A command entry point clears it before it runs
' and checks it after, so a fault caught + swallowed deep in a called module still surfaces to the
' user exactly once (see Command.ReportIfLogged). Informational logs (Number 0) never set it.
Private mbHadError As Boolean
' True while this instance posts to the Message Center: a HandleError reached from inside that
' posting (a faulting GetTranslation) only writes the .log.
Private mbPosting As Boolean

' === CLASS INITIALIZATION ===
Private Sub Class_Initialize()
    On Error Resume Next
    msLogFilePath = GetLogFilePath()
End Sub

' === PUBLIC METHODS ===
' Main error handling and logging method. Every write is also posted to the Message Center: an
' Error (translated) when this write creates the .log file, then a Warning (English) for the entry.
' Details: optional English context (values involved), logged on its own indented line.
' Returns: True if successfully logged, False otherwise
Public Function HandleError(ByVal Description As String, _
                          ByVal Number As Long, _
                          ByVal Source As String, _
                          Optional ByVal ModuleName As String = "", _
                          Optional ByVal Details As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim ErrorMessage As String
    Dim FileNum As Integer
    Dim bCreated As Boolean
    
    HandleError = False
    
    ' Build and format complete error message
    ErrorMessage = FormatErrorMessage(Description, Number, Source, ModuleName, Details)

    ' Flag real faults (not informational Number-0 logs) so a command entry point can report a
    ' swallowed-downstream fault to the user exactly once.
    If Number > 0 Then mbHadError = True
    
    ' Check log rotation (a rotation renames the file, so the append below creates a new one)
    RotateIfNeeded
    bCreated = (Len(Dir(msLogFilePath)) = 0)
    
    ' Write to log file
    FileNum = FreeFile
    Open msLogFilePath For Append As #FileNum
    Print #FileNum, ErrorMessage
    Close #FileNum
    
    HandleError = True

    If Not mbPosting Then
        mbPosting = True
        PostToMessageCenter ErrorMessage, bCreated, "Log: " & msLogFilePath
        mbPosting = False
    End If
    Exit Function

ErrorHandler:
    HandleError = False
    ' The append failed (read-only folder, locked file): the Warning still goes out, so an unloggable
    ' fault is not silent
    If FileNum > 0 Then CloseQuietly FileNum
    If Not mbPosting Then
        mbPosting = True
        PostToMessageCenter ErrorMessage, False, "Log write failed: " & msLogFilePath
        mbPosting = False
    End If
End Function

' True if a real fault (Number > 0) has been logged since the last ClearErrorFlag. Command entry
' points use this to surface a downstream fault that was caught + swallowed in a called module.
Public Property Get HadError() As Boolean
    HadError = mbHadError
End Property

' Reset the fault flag. A command entry point calls this before it runs so HadError then reflects
' only faults raised during that command's execution.
Public Sub ClearErrorFlag()
    mbHadError = False
End Sub

Public Sub UpdateLogFilePath()
    On Error Resume Next
    msLogFilePath = GetLogFilePath()
End Sub

' Clear log file
Public Sub ClearLogFile()
    On Error Resume Next
    
    Dim FileNum As Integer
    FileNum = FreeFile
    Open msLogFilePath For Output As #FileNum
    Close #FileNum
End Sub

' Get current log file path (for testing)
Public Property Get LogFilePath() As String
    LogFilePath = msLogFilePath
End Property

' Get last log entry (for testing)
Public Function GetLastLogEntry() As String
    On Error GoTo ErrorHandler
    
    Dim FileNum As Integer
    Dim Line As String
    Dim LastLine As String
    
    If Len(Dir(msLogFilePath)) = 0 Then
        GetLastLogEntry = ""
        Exit Function
    End If
    
    FileNum = FreeFile
    Open msLogFilePath For Input As #FileNum
    Do While Not EOF(FileNum)
        Line Input #FileNum, Line
        If Len(Line) > 0 Then LastLine = Line
    Loop
    Close #FileNum
    
    GetLastLogEntry = LastLine
    Exit Function
    
ErrorHandler:
    GetLastLogEntry = ""
End Function

' === PRIVATE METHODS ===
' Format complete error message with timestamp
Private Function FormatErrorMessage(ByVal Description As String, _
                                   ByVal Number As Long, _
                                   ByVal Source As String, _
                                   ByVal ModuleName As String, _
                                   ByVal Details As String) As String
    Dim Result As String
    
    ' Timestamp
    Result = "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "]"
    
    ' Module name if provided
    If Len(ModuleName) > 0 Then
        Result = Result & " [" & ModuleName & "]"
    End If
    
    ' Error details
    If Number > 0 Then
        Result = Result & " Error " & Number
        If Len(Source) > 0 Then
            Result = Result & " (" & Source & ")"
        End If
        Result = Result & ": "
    Else
        Result = Result & " "
    End If
    
    Result = Result & Description

    ' Append the current call-stack chain, if any is available and non-empty (see CallStackClass -
    ' read-only, tolerant of BootLoader.CallStack not existing yet/anymore, so a missing/torn-down
    ' stack never breaks logging). Own line, indented, so the primary message stays grep-stable.
    Dim sStack As String
    On Error Resume Next
    sStack = CallStack.Snapshot
    On Error GoTo 0
    If Len(sStack) > 0 Then
        Result = Result & vbCrLf & "    CallStack: " & sStack
    End If

    If Len(Details) > 0 Then
        ' Continuation lines indented under "Details: " so a multi-line value stays one block
        Result = Result & vbCrLf & "    Details: " & Replace(Details, vbCrLf, vbCrLf & "             ")
    End If

    FormatErrorMessage = Result
End Function

' Check and rotate log file if too large
Private Sub RotateIfNeeded()
    On Error Resume Next
    
    Dim FileNum As Integer
    Dim FileSize As Long
    
    If Len(Dir(msLogFilePath)) = 0 Then Exit Sub
    
    FileNum = FreeFile
    Open msLogFilePath For Binary Access Read As #FileNum
    FileSize = LOF(FileNum)
    Close #FileNum
    
    If FileSize > MAX_LOG_FILE_SIZE Then
        Dim BackupPath As String
        BackupPath = Replace(msLogFilePath, ".log", "-" & Format(Now, "yyyymmdd_hhnnss") & ".old")
        Name msLogFilePath As BackupPath
    End If
End Sub

' Get log file path
Private Function GetLogFilePath() As String
    On Error Resume Next
    
    Dim Path As String
    If ActiveDesignFile Is Nothing Then
        If Err.Number = -2147221504 Then Err.Clear
        If Not Application.VBE.ActiveVBProject Is Nothing Then
            Path = Application.VBE.ActiveVBProject.FileName
            If Len(Path) > 0 Then
                Path = Left(Path, InStrRev(Path, "\")) & "ARES - LogFile.log"
            End If
        End If
    Else
        Path = ActiveDesignFile.FullName & " - ARESLogFile.log"
    End If
    
    If Len(Path) = 0 Then
        Path = Environ("TEMP") & "\ARES - " & Format(Date, "yyyymmdd") & ".log"
    End If
    
    GetLogFilePath = Path
End Function

' Posts the entry to the Message Center: the Error first when the write created the .log, then the
' Warning, so the newest message (GetMessage(0)) is the Warning. LogNote ends the Warning's Details
' ("Log: <path>", or "Log write failed: <path>"). Never raises.
Private Sub PostToMessageCenter(ByVal Entry As String, ByVal Created As Boolean, ByVal LogNote As String)
    On Error Resume Next

    Dim sText As String
    Dim nPos As Long

    If Created Then
        If LangManager.IsInit Then
            sText = LangManager.GetTranslation("LogFileCreated", msLogFilePath)
        End If
        If Len(sText) = 0 Then sText = Replace(LOG_FILE_CREATED_EN, "{0}", msLogFilePath)
        Messaging.PostMessage sText, msLogFilePath, msdMessageCenterPriorityError
    End If

    ' Warning text = first line of the entry, without its "[timestamp] " prefix
    sText = Entry
    nPos = InStr(sText, vbCrLf)
    If nPos > 0 Then sText = Left(sText, nPos - 1)
    nPos = InStr(sText, "] ")
    If nPos > 0 Then sText = Mid(sText, nPos + 2)
    Messaging.PostMessage "ARES: " & sText, Entry & vbCrLf & LogNote, msdMessageCenterPriorityWarning
End Sub

' Closes a file number without raising: used from HandleError's error label, where a second fault
' would escape to the caller
Private Sub CloseQuietly(ByVal FileNum As Integer)
    On Error Resume Next
    Close #FileNum
End Sub
