' Module: Config
' Description: This module provides functions to manage configuration variables in MicroStation with silent error handling.
' It includes functions to get, set, configuration variables and RemoveValue, ensuring that operations are performed
' without interrupting the workflow in case of errors.
' Delete a configuration variable is possible but not saved if you restart MS, use RemoveValue instead.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfigClass, ARESConstants ,ErrorHandlerClass
Option Explicit

' Function to get the value of a configuration variable
' Returns ARESConstants.ARES_NAVD if configuration variable is not defined
Public Function GetVar(ByVal StrKey As String) As String
    On Error GoTo ErrorHandler

    ' Validate input parameter
    If Not IsValidKey(StrKey) Then
        GetVar = ARESConstants.ARES_NAVD
        Exit Function
    End If

    ' Initialize the return value
    GetVar = ARESConstants.ARES_NAVD

    ' Check if the configuration variable is defined
    If Application.ActiveWorkspace.IsConfigurationVariableDefined(StrKey) Then
        ' Retrieve the value of the configuration variable
        GetVar = Application.ActiveWorkspace.ConfigurationVariableValue(StrKey)
    End If

    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Config.GetVar"
    ' In case of error, return ARES_NAVD constant
    GetVar = ARESConstants.ARES_NAVD
End Function

' Function to set the value of a configuration variable
' Creates the variable and sets the definition. If the configuration is already defined, it just updates the definition.
Public Function SetVar(ByVal StrKey As String, ByVal strValue As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Initialize the return value
    SetVar = False
    
    ' Validate input parameters
    If Not IsValidKey(StrKey) Then Exit Function
    If strValue = vbNullString Then strValue = ""
    
    ' Set the configuration variable value
    Application.ActiveWorkspace.AddConfigurationVariable StrKey, strValue, True
    SetVar = True
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Config.SetVar"
    ' In case of error, return False
    SetVar = False
End Function

' Function to remove a value of a configuration variable
Public Function RemoveValue(ByVal StrKey As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Initialize the return value
    RemoveValue = False
    
    ' Validate input parameter
    If Not IsValidKey(StrKey) Then Exit Function
    
    ' Check if the configuration variable is defined
    If Application.ActiveWorkspace.IsConfigurationVariableDefined(StrKey) Then
        ' Remove the configuration variable value by setting it to an empty string
        RemoveValue = SetVar(StrKey, "")
    End If
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Config.RemoveValue"
    ' In case of error, return False
    RemoveValue = False
End Function

' Function to check if a key is valid
' Returns True if key is not empty, null, or only whitespace
Private Function IsValidKey(ByVal StrKey As String) As Boolean
    On Error GoTo ErrorHandler
    
    IsValidKey = False
    
    ' Check for null, empty, or whitespace-only strings
    If StrKey <> vbNullString And Len(Trim(StrKey)) > 0 Then
        IsValidKey = True
    End If
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "Config.IsValidKey"
    IsValidKey = False
End Function