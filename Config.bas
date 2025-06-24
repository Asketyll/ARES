' Module: Config
' Description: This module provides functions to manage configuration variables in MicroStation with silent error handling.
' It includes functions to get, set, configuration variables and RemoveValue, ensuring that operations are performed
' without interrupting the workflow in case of errors.
' Delete a configuration variable is possible but not saved if you restart MS, use RemoveValue instead.

Option Explicit

' Function to get the value of a configuration variable
' Returne ARES_VAR.ARES_NAVD if configuration variable is not defined
Public Function GetVar(ByVal key As String) As String
    On Error GoTo ErrorHandler

    ' Initialize the return value
    GetVar = ""

    ' Check if the configuration variable is defined
    If Application.ActiveWorkspace.IsConfigurationVariableDefined(key) Then
        ' Retrieve the value of the configuration variable
        GetVar = Application.ActiveWorkspace.ConfigurationVariableValue(key)
    Else
        GetVar = ARES_VAR.ARES_NAVD
    End If

    Exit Function

ErrorHandler:
    GetVar = ""
End Function

' Function to set the value of a configuration variable
' creates the variable and sets the definition. If the configuration is already defined, it just updates the definition.
Public Function SetVar(ByVal key As String, ByVal Value As String) As Boolean
    On Error GoTo ErrorHandler

    ' Initialize the return value
    SetVar = False

    Application.ActiveWorkspace.AddConfigurationVariable key, Value, True
    SetVar = True

    Exit Function

ErrorHandler:
    SetVar = False
End Function

' Function to Remove a value of a configuration variable
Public Function RemoveValue(ByVal key As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Initialize the return value
    RemoveValue = False
    
    ' Check if the configuration variable is defined
    If Application.ActiveWorkspace.IsConfigurationVariableDefined(key) Then
        ' Remove the configuration variable value
        If SetVar(key, "") Then RemoveValue = True
    End If

    Exit Function

ErrorHandler:
    RemoveValue = False
End Function
