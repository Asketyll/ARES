'Class Module: ErrorHandlerClass
' Description: This module Handles error logging and management.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: None
Option Explicit

' Enum for log levels
Public Enum LogLevel
    LOG_LEVEL_INFO
    LOG_LEVEL_WARNING
    LOG_LEVEL_ERROR
End Enum

' Constants for log file path and settings
Private Const MAX_LOG_FILE_SIZE As Long = 1048576 ' 1 MB
Private logFilePath As String

' Function to log a message with a specified log level and module name
Public Sub LogMessage(ByVal message As String, ByVal level As LogLevel, Optional ByVal moduleName As String = "")
    On Error Resume Next ' Do not interrupt execution in case of logging error
    Dim logFile As Integer
    Dim LogMessage As String
    Dim logFilePath As String
    Dim vbProj As Object
    
    If ActiveDesignFile Is Nothing Then
        Set vbProj = Application.VBE.ActiveVBProject
        logFilePath = vbProj.FileName
        logFilePath = Left(logFilePath, InStrRev(logFilePath, "\") - 1)
        logFilePath = logFilePath & "\DefaultLogFile.txt"
    Else
        ' Define the path to the log file
        logFilePath = ActiveDesignFile.FullName & " - LogFile.txt"
    End If
    
    ' Build the log message
    LogMessage = "[" & Now & "] " & "[" & GetLogLevelString(level) & "]"
    If moduleName <> "" Then
        LogMessage = LogMessage & " [" & moduleName & "]"
    End If
    LogMessage = LogMessage & " " & message
    
    ' Check log file size and rename if it is too large
    CheckLogFileSize

    ' Write to the log file
    logFile = FreeFile
    Open logFilePath For Append As #logFile
    Print #logFile, LogMessage
    Close #logFile
    
End Sub

' Method to log an info message
Public Sub LogInfo(ByVal message As String, Optional ByVal moduleName As String = "")
    LogMessage message, LOG_LEVEL_INFO, moduleName
End Sub

' Method to log a warning message
Public Sub LogWarning(ByVal message As String, Optional ByVal moduleName As String = "")
    LogMessage message, LOG_LEVEL_WARNING, moduleName
End Sub

' Method to log an error message
Public Sub LogError(ByVal message As String, Optional ByVal moduleName As String = "")
    LogMessage message, LOG_LEVEL_ERROR, moduleName
End Sub

' Method to get the string representation of a log level
Private Function GetLogLevelString(ByVal level As LogLevel) As String
    Select Case level
        Case LOG_LEVEL_INFO
            GetLogLevelString = "INFO"
        Case LOG_LEVEL_WARNING
            GetLogLevelString = "WARNING"
        Case LOG_LEVEL_ERROR
            GetLogLevelString = "ERROR"
        Case Else
            GetLogLevelString = "UNKNOWN"
    End Select
End Function

' Method to clear the log file
Public Sub ClearLogFile()
    On Error Resume Next
    Dim logFile As Integer

    logFile = FreeFile
    Open LOG_FILE_PATH For Output As #logFile
    Close #logFile
End Sub

' Method to check the log file size and rotate if necessary
Private Sub CheckLogFileSize()
    Dim fileSize As Long
    Dim fileNum As Integer
    
    ' Check if the log file exists
    If Dir(logFilePath) <> "" Then
        fileNum = FreeFile
        Open logFilePath For Binary Access Read As #fileNum
        fileSize = LOF(fileNum)
        Close #fileNum

        ' If the file is too large, rename it
        If fileSize > MAX_LOG_FILE_SIZE Then
            Dim backupPath As String
            Dim i As Integer
            i = 1
            backupPath = logFilePath & "." & i
            Do While Dir(backupPath) <> ""
                i = i + 1
                backupPath = logFilePath & "." & i
            Loop
            Name logFilePath As backupPath
        End If
    End If
End Sub

' Method to handle an error and log it
Public Sub HandleError(ByVal errDescription As String, ByVal errNumber As Long, ByVal errSource As String, Optional ByVal moduleName As String = "")
    Dim errorMessage As String
    errorMessage = "Error " & errNumber & " (" & errSource & "): " & errDescription
    LogError errorMessage, moduleName
End Sub
