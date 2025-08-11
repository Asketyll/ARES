' Module: Config
' Description: This module provides functions to manage configuration variables in MicroStation with silent error handling.
' It includes functions to get, set, configuration variables and RemoveValue, ensuring that operations are performed
' without interrupting the workflow in case of errors.
' Delete a configuration variable is possible but not saved if you restart MS, use RemoveValue instead.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ARESConfig, ARESConstants ,ErrorHandlerClass
Option Explicit

' Function to get the value of a configuration variable
' Returne ARES_VAR.ARES_NAVD if configuration variable is not defined
Public Function GetVar(ByVal key As String) As String
    On Error GoTo ErrorHandler

    ' Initialize the return value
    GetVar = ARES_NAVD

    ' Check if the configuration variable is defined
    If Application.ActiveWorkspace.IsConfigurationVariableDefined(key) Then
        ' Retrieve the value of the configuration variable
        GetVar = Application.ActiveWorkspace.ConfigurationVariableValue(key)
    End If

    Exit Function

ErrorHandler:
    ErrorHandler.LogError Err.Description, "Config.GetVar"
    ' In case of error, return an empty string
    GetVar = ""
End Function

' Function to set the value of a configuration variable
' creates the variable and sets the definition. If the configuration is already defined, it just updates the definition.
Public Function SetVar(ByVal key As String, ByVal Value As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Initialize the return value
    SetVar = False
    
    If Not IsValidKey(key) Then Exit Function
    
    ' Set the configuration variable value
    Application.ActiveWorkspace.AddConfigurationVariable key, Value, True
    SetVar = True
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.LogError Err.Description, "Config.SetVar"
    ' In case of error, return False
    SetVar = False
End Function

' Function to Remove a value of a configuration variable
Public Function RemoveValue(ByVal key As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Initialize the return value
    RemoveValue = False
    
    ' Check if the configuration variable is defined
    If Application.ActiveWorkspace.IsConfigurationVariableDefined(key) Then
        ' Remove the configuration variable value by setting it to an empty string
        RemoveValue = SetVar(key, "")
    End If
    
    Exit Function
    
ErrorHandler:
    ErrorHandler.LogError Err.Description, "Config.RemoveValue"
    ' In case of error, return False
    RemoveValue = False
End Function

' Function to check if a key is valid
Private Function IsValidKey(ByVal key As String) As Boolean
    IsValidKey = Not (key = "") Or (key = " ")
End Function
