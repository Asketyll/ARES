' Module: Config
' Description: This module provides functions to manage configuration variables in MicroStation with silent error handling.
' It includes functions to get, set, configuration variables and RemoveValue, ensuring that operations are performed
' without interrupting the workflow in case of errors.
' Delete a configuration variable is possible but not saved if you restart MS, use RemoveValue.

Option Explicit

' Function to get the value of a configuration variable
Public Function GetVar(ByVal Key As String) As String
    On Error GoTo ErrorHandler

    ' Initialize the return value
    GetVar = ""

    ' Check if the configuration variable is defined
    If Application.ActiveWorkspace.IsConfigurationVariableDefined(Key) Then
        ' Retrieve the value of the configuration variable
        GetVar = Application.ActiveWorkspace.ConfigurationVariableValue(Key)
    End If

    Exit Function

ErrorHandler:
    GetVar = ""
End Function

' Function to set the value of a configuration variable
' creates the variable and sets the definition. If the configuration is already defined, it just updates the definition.
Public Function SetVar(ByVal Key As String, ByVal Value As String) As Boolean
    On Error GoTo ErrorHandler

    ' Initialize the return value
    SetVar = False

    Application.ActiveWorkspace.AddConfigurationVariable Key, Value, True
    SetVar = True

    Exit Function

ErrorHandler:
    SetVar = False
End Function

' Function to Remove a value of a configuration variable
Public Function RemoveValue(ByVal Key As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Initialize the return value
    RemoveValue = False
    
    ' Check if the configuration variable is defined
    If Application.ActiveWorkspace.IsConfigurationVariableDefined(Key) Then
        ' Remove the configuration variable value
        If SetVar(Key, "") Then RemoveValue = True
    End If

    Exit Function

ErrorHandler:
    RemoveValue = False
End Function
