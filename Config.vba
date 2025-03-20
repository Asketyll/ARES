' Module: Config
' Description: This module provides functions to manage configuration variables in MicroStation with silent error handling.
' It includes functions to get, set, and delete configuration variables, ensuring that operations are performed
' without interrupting the workflow in case of errors. The module is designed to maintain the integrity of
' configuration variables by checking their existence before performing operations and handling errors.

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
Public Function SetVar(ByVal Key As String, ByVal Value As String) As Boolean
    On Error GoTo ErrorHandler

    ' Initialize the return value
    SetVar = False

    ' Check if the configuration variable is defined
    If Application.ActiveWorkspace.IsConfigurationVariableDefined(Key) Then
        ' Replace the existing variable
        SetVar = ReplaceVar(Key, Value)
    Else
        ' Create a new variable
        SetVar = CreateVar(Key, Value)
    End If

    Exit Function

ErrorHandler:
    SetVar = False
End Function

' Function to delete a configuration variable
Public Function DeleteVar(ByVal Key As String) As Boolean
    On Error GoTo ErrorHandler

    ' Initialize the return value
    DeleteVar = False

    ' Check if the configuration variable is defined
    If Application.ActiveWorkspace.IsConfigurationVariableDefined(Key) Then
        ' Remove the configuration variable
        Application.ActiveWorkspace.RemoveConfigurationVariable Key
        DeleteVar = True
    End If

    Exit Function

ErrorHandler:
    DeleteVar = False
End Function

' Private function to create a new configuration variable
Private Function CreateVar(ByVal Key As String, ByVal Value As String) As Boolean
    On Error GoTo ErrorHandler

    ' Initialize the return value
    CreateVar = False

    ' Check if the variable does not exist and the value is not empty
    If Not Application.ActiveWorkspace.IsConfigurationVariableDefined(Key) And Value <> "" Then
        ' Add the new configuration variable
        Application.ActiveWorkspace.AddConfigurationVariable Key, Value
        CreateVar = True
    End If

    Exit Function

ErrorHandler:
    CreateVar = False
End Function

' Private function to replace the value of an existing configuration variable
Private Function ReplaceVar(ByVal Key As String, ByVal Value As String) As Boolean
    On Error GoTo ErrorHandler

    ' Initialize the return value
    ReplaceVar = False

    ' Delete the existing variable and create a new one with the same key
    If DeleteVar(Key) Then
        ReplaceVar = CreateVar(Key, Value)
    End If

    Exit Function

ErrorHandler:
    ReplaceVar = False
End Function
