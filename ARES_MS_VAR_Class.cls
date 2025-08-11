' Class Module: ARES_MS_VAR_Class
' Description: This class module represents a configuration variable in MicroStation.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: None
Option Explicit

' Property to store the key of the configuration variable
Public key As String
' Property to store the value of the configuration variable
Private pValue As String
' Property to store the default value of the configuration variable
Public Default As String

' Property to get or set the value of the configuration variable
Public Property Get Value() As String
    Value = pValue
End Property

Public Property Let Value(ByVal newValue As String)
    Dim oldValue As String
    oldValue = pValue
    pValue = newValue
    ' Optional: trigger an event when the value changes
    ' OnValueChanged oldValue, newValue
End Property

' Method to reset the value to its default
Public Sub ResetToDefault()
    pValue = Default
End Sub

' Method to check if the current value is different from the default
Public Function IsDefault() As Boolean
    IsDefault = (pValue = Default)
End Function
