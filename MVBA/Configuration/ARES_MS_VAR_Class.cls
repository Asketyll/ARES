' Class Module: ARES_MS_VAR_Class
' Description: Represents a MicroStation configuration variable with key, value, and default handling
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ErrorHandlerClass (via BootLoader)
Option Explicit

' === PUBLIC PROPERTIES ===
Public Key As String                ' Configuration variable key
Public DefaultValue As String       ' Default value for reset operations

' === PRIVATE MEMBERS ===
Private mstrValue As String         ' Current value of the configuration variable
Private mbIsModified As Boolean     ' Flag to track if value has been modified

' Property to get the current value of the configuration variable
Public Property Get Value() As String
    On Error GoTo ErrorHandler
    Value = mstrValue
    Exit Property
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARES_MS_VAR_Class.Value_Get"
    Value = DefaultValue
End Property

' Property to set the value of the configuration variable
Public Property Let Value(ByVal strNewValue As String)
    On Error GoTo ErrorHandler
    
    Dim strOldValue As String
    strOldValue = mstrValue
    
    ' Update the value
    mstrValue = strNewValue
    
    ' Track modification state
    mbIsModified = (mstrValue <> DefaultValue)

    ' Persist to MicroStation configuration
    If Len(Key) > 0 Then
        Config.SetVar Key, strNewValue
    End If
	
    Exit Property
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARES_MS_VAR_Class.Value_Let"
End Property

' Reset the value to its default value
Public Sub ResetToDefault()
    On Error GoTo ErrorHandler
    
    mstrValue = DefaultValue
    mbIsModified = False
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARES_MS_VAR_Class.ResetToDefault"
End Sub

' Check if the current value equals the default value
Public Function IsDefault() As Boolean
    On Error GoTo ErrorHandler
    
    IsDefault = (mstrValue = DefaultValue)
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARES_MS_VAR_Class.IsDefault"
    IsDefault = True ' Assume default in case of error
End Function

' Check if the value has been modified from default
Public Function IsModified() As Boolean
    On Error GoTo ErrorHandler
    
    IsModified = mbIsModified
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARES_MS_VAR_Class.IsModified"
    IsModified = False
End Function

' Validate if the current value is not empty or null
Public Function IsValid() As Boolean
    On Error GoTo ErrorHandler
    
    IsValid = (Len(Trim(mstrValue)) > 0)
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARES_MS_VAR_Class.IsValid"
    IsValid = False
End Function

' Initialize the variable with key and default value
Public Sub Initialize(ByVal StrKey As String, ByVal strDefaultValue As String)
    On Error GoTo ErrorHandler
    
    Key = StrKey
    DefaultValue = strDefaultValue
    mstrValue = strDefaultValue
    mbIsModified = False
    
    Exit Sub
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARES_MS_VAR_Class.Initialize"
End Sub

' Get a string representation of this variable for debugging
Public Function ToString() As String
    On Error GoTo ErrorHandler
    
    ToString = "Key: " & Key & ", Value: " & mstrValue & ", Default: " & DefaultValue & ", Modified: " & CStr(mbIsModified)
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ARES_MS_VAR_Class.ToString"
    ToString = "[Error getting variable info]"
End Function

' Clean up when object is destroyed
Private Sub Class_Terminate()
    On Error Resume Next
    Key = ""
    DefaultValue = ""
    mstrValue = ""
    mbIsModified = False
End Sub