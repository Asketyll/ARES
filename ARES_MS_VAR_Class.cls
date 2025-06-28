' Class Module: ARES_MS_VAR_Class
' Description: This class module represents a configuration variable in MicroStation.
' It encapsulates the properties and behaviors of a configuration variable, including its key, value, and default value.
' License: This project is licensed under the AGPL-3.0.

Option Explicit

' Property to store the key of the configuration variable
Public key As String

' Property to store the value of the configuration variable
Public Value As String

' Property to store the default value of the configuration variable
Public Default As String

' Initialize a new instance of the ARES_MS_VAR_Class
Public Sub Initialize(ByVal varKey As String, ByVal varValue As String, ByVal varDefault As String)
    On Error GoTo ErrorHandler

    ' Set the key, value, and default value of the configuration variable
    key = varKey
    Value = varValue
    Default = varDefault

    Exit Sub

ErrorHandler:
    ' Handle any errors silently
    Err.Clear
End Sub

' Reset the value of the configuration variable to its default value
Public Sub ResetToDefault()
    On Error GoTo ErrorHandler

    ' Reset the value to the default value
    Value = Default

    Exit Sub

ErrorHandler:
    ' Handle any errors silently
    Err.Clear
End Sub
