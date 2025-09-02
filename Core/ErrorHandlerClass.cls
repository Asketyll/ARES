' Class Module: ErrorHandlerClass
' Description: Centralized error handling and logging system for ARES application
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConstants
Option Explicit

' Enum for log levels
Public Enum logLevel
    LOG_LEVEL_DEBUG = 0
    LOG_LEVEL_INFO = 1
    LOG_LEVEL_WARNING = 2
    LOG_LEVEL_ERROR = 3
    LOG_LEVEL_CRITICAL = 4
End Enum

' === CONSTANTS ===
Private Const MAX_LOG_FILE_SIZE As Long = 1048576    ' 1 MB
Private Const LOG_FILE_EXTENSION As String = ".log"
Private Const BACKUP_EXTENSION As String = ".old"

' === PRIVATE MEMBERS ===
Private mstrLogFilePath As String
Private mbLoggingEnabled As Boolean
Private mCurrentLogLevel As logLevel

' Initialize the error handler with logging preferences
Public Sub Initialize(Optional ByVal bEnableLogging As Boolean = True, Optional ByVal logLevel As logLevel = LOG_LEVEL_INFO)
    On Error Resume Next
    
    mbLoggingEnabled = bEnableLogging
    mCurrentLogLevel = logLevel
    
    ' Set up log file path
    mstrLogFilePath = GetLogFilePath()
    
    ' Log initialization
    If mbLoggingEnabled Then
        WriteToLog "ErrorHandler initialized", LOG_LEVEL_INFO, "ErrorHandlerClass.Initialize"
    End If
End Sub

' Main logging method with level filtering
Public Sub LogMessage(ByVal strMessage As String, ByVal level As logLevel, Optional ByVal strModuleName As String = "")
    On Error Resume Next ' Do not interrupt execution in case of logging error
    
    ' Check if logging is enabled and level is appropriate
    If Not mbLoggingEnabled Or level < mCurrentLogLevel Then
        Exit Sub
    End If
    
    WriteToLog strMessage, level, strModuleName
End Sub

' Internal method to write to log file
Private Sub WriteToLog(ByVal strMessage As String, ByVal level As logLevel, ByVal strModuleName As String)
    On Error Resume Next
    
    Dim intLogFile As Integer
    Dim strFormattedMessage As String
    Dim strCurrentLogPath As String
    
    ' Get current log file path
    strCurrentLogPath = GetLogFilePath()
    
    ' Build the formatted log message
    strFormattedMessage = "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "] " & _
                         "[" & GetLogLevelString(level) & "]"
                         
    If Len(strModuleName) > 0 Then
        strFormattedMessage = strFormattedMessage & " [" & strModuleName & "]"
    End If
    
    strFormattedMessage = strFormattedMessage & " " & strMessage
    
    ' Check log file size before writing
    CheckAndRotateLogFile strCurrentLogPath

    ' Write to the log file
    intLogFile = FreeFile
    Open strCurrentLogPath For Append As #intLogFile
    Print #intLogFile, strFormattedMessage
    Close #intLogFile
End Sub

' Convenience methods for different log levels
Public Sub LogDebug(ByVal strMessage As String, Optional ByVal strModuleName As String = "")
    LogMessage strMessage, LOG_LEVEL_DEBUG, strModuleName
End Sub

Public Sub LogInfo(ByVal strMessage As String, Optional ByVal strModuleName As String = "")
    LogMessage strMessage, LOG_LEVEL_INFO, strModuleName
End Sub

Public Sub LogWarning(ByVal strMessage As String, Optional ByVal strModuleName As String = "")
    LogMessage strMessage, LOG_LEVEL_WARNING, strModuleName
End Sub

Public Sub LogError(ByVal strMessage As String, Optional ByVal strModuleName As String = "")
    LogMessage strMessage, LOG_LEVEL_ERROR, strModuleName
End Sub

Public Sub LogCritical(ByVal strMessage As String, Optional ByVal strModuleName As String = "")
    LogMessage strMessage, LOG_LEVEL_CRITICAL, strModuleName
End Sub

' Get string representation of log level
Private Function GetLogLevelString(ByVal level As logLevel) As String
    Select Case level
        Case LOG_LEVEL_DEBUG
            GetLogLevelString = "DEBUG"
        Case LOG_LEVEL_INFO
            GetLogLevelString = "INFO"
        Case LOG_LEVEL_WARNING
            GetLogLevelString = "WARN"
        Case LOG_LEVEL_ERROR
            GetLogLevelString = "ERROR"
        Case LOG_LEVEL_CRITICAL
            GetLogLevelString = "CRIT"
        Case Else
            GetLogLevelString = "UNKNOWN"
    End Select
End Function

' Clear the current log file
Public Sub ClearLogFile()
    On Error Resume Next
    Dim intLogFile As Integer
    Dim strCurrentLogPath As String
    
    strCurrentLogPath = GetLogFilePath()
    
    intLogFile = FreeFile
    Open strCurrentLogPath For Output As #intLogFile
    Close #intLogFile
    
    If mbLoggingEnabled Then
        WriteToLog "Log file cleared", LOG_LEVEL_INFO, "ErrorHandlerClass.ClearLogFile"
    End If
End Sub

' Check log file size and rotate if necessary
Private Sub CheckAndRotateLogFile(ByVal strLogPath As String)
    On Error Resume Next
    
    Dim lngFileSize As Long
    Dim intFileNum As Integer
    
    ' Check if the log file exists and get its size
    If Len(Dir(strLogPath)) > 0 Then
        intFileNum = FreeFile
        Open strLogPath For Binary Access Read As #intFileNum
        lngFileSize = LOF(intFileNum)
        Close #intFileNum

        ' If the file is too large, rotate it
        If lngFileSize > MAX_LOG_FILE_SIZE Then
            RotateLogFile strLogPath
        End If
    End If
End Sub

' Rotate log file by renaming to backup
Private Sub RotateLogFile(ByVal strLogPath As String)
    On Error Resume Next
    
    Dim strBackupPath As String
    Dim intCounter As Integer
    
    ' Create backup file name with timestamp
    strBackupPath = Left(strLogPath, InStrRev(strLogPath, ".") - 1) & _
                   "_" & Format(Now, "yyyymmdd_hhnnss") & BACKUP_EXTENSION
    
    ' If backup already exists, add counter
    intCounter = 1
    Do While Len(Dir(strBackupPath)) > 0
        strBackupPath = Left(strLogPath, InStrRev(strLogPath, ".") - 1) & _
                       "_" & Format(Now, "yyyymmdd_hhnnss") & "_" & intCounter & BACKUP_EXTENSION
        intCounter = intCounter + 1
    Loop
    
    ' Rename current log to backup
    Name strLogPath As strBackupPath
End Sub

' Main error handling method with flexible severity detection
Public Sub HandleError(ByVal strErrorDescription As String, _
                      Optional ByVal lngErrorNumber As Long = 0, _
                      Optional ByVal strErrorSource As String = "", _
                      Optional ByVal strModuleName As String = "", _
                      Optional ByVal strSeverity As String = "")
    On Error Resume Next
    
    Dim strErrorMessage As String
    Dim logLevel As logLevel
    
    ' Determine log level based on severity or error number
    logLevel = DetermineLogLevel(lngErrorNumber, strSeverity)
    
    ' Build comprehensive error message
    strErrorMessage = BuildErrorMessage(strErrorDescription, lngErrorNumber, strErrorSource)
    
    ' Log the error
    LogMessage strErrorMessage, logLevel, strModuleName
    
    ' For critical errors, also show message to user in development mode
    If logLevel = LOG_LEVEL_CRITICAL And Application.VBE.ActiveVBProject.Mode = vbext_vm_Design Then
        MsgBox "Critical Error in ARES: " & strErrorMessage, vbCritical + vbOKOnly, "ARES Error Handler"
    End If
End Sub

' Determine appropriate log level based on error info
Private Function DetermineLogLevel(ByVal lngErrorNumber As Long, ByVal strSeverity As String) As logLevel
    On Error Resume Next
    
    ' Check explicit severity first
    Select Case UCase(Trim(strSeverity))
        Case "DEBUG"
            DetermineLogLevel = LOG_LEVEL_DEBUG
        Case "INFO"
            DetermineLogLevel = LOG_LEVEL_INFO
        Case "WARNING", "WARN"
            DetermineLogLevel = LOG_LEVEL_WARNING
        Case "ERROR"
            DetermineLogLevel = LOG_LEVEL_ERROR
        Case "CRITICAL", "CRIT"
            DetermineLogLevel = LOG_LEVEL_CRITICAL
        Case Else
            ' Determine by error number if no explicit severity
            If lngErrorNumber = 0 Then
                DetermineLogLevel = LOG_LEVEL_INFO
            ElseIf lngErrorNumber >= 1000 Then
                DetermineLogLevel = LOG_LEVEL_CRITICAL
            ElseIf lngErrorNumber >= 100 Then
                DetermineLogLevel = LOG_LEVEL_ERROR
            Else
                DetermineLogLevel = LOG_LEVEL_WARNING
            End If
    End Select
End Function

' Build formatted error message
Private Function BuildErrorMessage(ByVal strDescription As String, ByVal lngNumber As Long, ByVal strSource As String) As String
    On Error Resume Next
    
    Dim strResult As String
    
    If lngNumber > 0 Then
        strResult = "Error " & lngNumber
    Else
        strResult = "Event"
    End If
    
    If Len(strSource) > 0 Then
        strResult = strResult & " (" & strSource & ")"
    End If
    
    strResult = strResult & ": " & strDescription
    
    BuildErrorMessage = strResult
End Function

' Get appropriate log file path
Private Function GetLogFilePath() As String
    On Error Resume Next
    
    Dim strPath As String
    
    If Len(mstrLogFilePath) > 0 Then
        GetLogFilePath = mstrLogFilePath
        Exit Function
    End If
    
    If Not ActiveDesignFile Is Nothing Then
        ' Use design file path
        strPath = ActiveDesignFile.FullName & "_ARES" & LOG_FILE_EXTENSION
    Else
        ' Use VBA project path as fallback
        Dim oVBProj As Object
        Set oVBProj = Application.VBE.ActiveVBProject
        If Not oVBProj Is Nothing And Len(oVBProj.FileName) > 0 Then
            strPath = Left(oVBProj.FileName, InStrRev(oVBProj.FileName, "\")) & "ARES_Default" & LOG_FILE_EXTENSION
        Else
            ' Last resort: temp directory
            strPath = Environ("TEMP") & "\ARES_" & Format(Now, "yyyymmdd") & LOG_FILE_EXTENSION
        End If
    End If
    
    mstrLogFilePath = strPath
    GetLogFilePath = strPath
End Function

' Clean up when class is destroyed
Private Sub Class_Terminate()
    On Error Resume Next
    
    If mbLoggingEnabled Then
        WriteToLog "ErrorHandler terminating", LOG_LEVEL_INFO, "ErrorHandlerClass.Class_Terminate"
    End If
    
    mstrLogFilePath = ""
    mbLoggingEnabled = False
End Sub