' Class Module: ErrorHandlerClass
' Description: Centralized error handling and logging for ARES application
' License: This project is licensed under the AGPL-3.0.
' Dependencies: None
Option Explicit

' === CONSTANTS ===
Private Const MAX_LOG_FILE_SIZE As Long = 1048576    ' 1 MB

' === PRIVATE MEMBERS ===
Private mLogFilePath As String

' === CLASS INITIALIZATION ===
Private Sub Class_Initialize()
    On Error Resume Next
    mLogFilePath = GetLogFilePath()
End Sub

' === PUBLIC METHODS ===
' Main error handling and logging method
' Returns: True if successfully logged, False otherwise
Public Function HandleError(ByVal Description As String, _
                          ByVal Number As Long, _
                          ByVal Source As String, _
                          Optional ByVal ModuleName As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim ErrorMessage As String
    Dim FileNum As Integer
    
    HandleError = False
    
    ' Build and format complete error message
    ErrorMessage = FormatErrorMessage(Description, Number, Source, ModuleName)
    
    ' Check log rotation
    RotateIfNeeded
    
    ' Write to log file
    FileNum = FreeFile
    Open mLogFilePath For Append As #FileNum
    Print #FileNum, ErrorMessage
    Close #FileNum
    
    ' Show to user in debug mode
    If Application.VBE.ActiveVBProject.Mode = vbext_vm_Design Then
        ShowError Description, Number, Source
    End If
    
    HandleError = True
    Exit Function
    
ErrorHandler:
    HandleError = False
End Function

Public Sub UpdateLogFilePath()
    On Error Resume Next
    mLogFilePath = GetLogFilePath()
End Sub

' Clear log file
Public Sub ClearLogFile()
    On Error Resume Next
    
    Dim FileNum As Integer
    FileNum = FreeFile
    Open mLogFilePath For Output As #FileNum
    Close #FileNum
End Sub

' Get current log file path (for testing)
Public Property Get LogFilePath() As String
    LogFilePath = mLogFilePath
End Property

' Get last log entry (for testing)
Public Function GetLastLogEntry() As String
    On Error GoTo ErrorHandler
    
    Dim FileNum As Integer
    Dim Line As String
    Dim LastLine As String
    
    If Len(Dir(mLogFilePath)) = 0 Then
        GetLastLogEntry = ""
        Exit Function
    End If
    
    FileNum = FreeFile
    Open mLogFilePath For Input As #FileNum
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
                                   ByVal ModuleName As String) As String
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
    FormatErrorMessage = Result
End Function

' Check and rotate log file if too large
Private Sub RotateIfNeeded()
    On Error Resume Next
    
    Dim FileNum As Integer
    Dim FileSize As Long
    
    If Len(Dir(mLogFilePath)) = 0 Then Exit Sub
    
    FileNum = FreeFile
    Open mLogFilePath For Binary Access Read As #FileNum
    FileSize = LOF(FileNum)
    Close #FileNum
    
    If FileSize > MAX_LOG_FILE_SIZE Then
        Dim BackupPath As String
        BackupPath = Replace(mLogFilePath, ".log", "-" & Format(Now, "yyyymmdd_hhnnss") & ".old")
        Name mLogFilePath As BackupPath
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

' Show error to user in debug mode
Private Sub ShowError(ByVal Description As String, ByVal Number As Long, ByVal Source As String)
    On Error Resume Next
    
    Dim Message As String
    
    If Number > 0 Then
        Message = "Error " & Number
        If Len(Source) > 0 Then Message = Message & " in " & Source
        Message = Message & vbCrLf & vbCrLf & Description
    Else
        Message = Description
    End If
    
    MsgBox Message, vbCritical + vbOKOnly, "ARES Error"
End Sub