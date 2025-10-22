' Module: LicenseManager
' Description: Manages license validation using AresLicenseValidator.dll COM component
' License: This project is licensed under the AGPL-3.0.
' Dependencies: ErrorHandlerClass
Option Explicit

' === PRIVATE MEMBERS ===
Private moLicenseValidator As Object
Private mbLicenseValid As Boolean
Private mstrLastError As String

' === PUBLIC PROPERTIES ===
Public Property Get IsLicenseValid() As Boolean
    IsLicenseValid = mbLicenseValid
End Property

Public Property Get LastError() As String
    LastError = mstrLastError
End Property

' Initialize and validate license
' Returns: True if license is valid, False otherwise
Public Function ValidateLicense() As Boolean
    On Error GoTo ErrorHandler
    
    ValidateLicense = False
    mbLicenseValid = False
    mstrLastError = ""
    
    ' Create COM instance of license validator
    If Not CreateValidatorInstance() Then
        Exit Function
    End If
    
    ' Validate the license
    If moLicenseValidator.ValidateLicense() Then
        mbLicenseValid = True
        ValidateLicense = True
        
    Else
        ' Get error details
        mstrLastError = moLicenseValidator.GetLastError()
        
        ' Log failure
        If Not ErrorHandler Is Nothing Then
            ErrorHandler.HandleError "License validation failed: " & mstrLastError, 0, "LicenseManager.ValidateLicense", "ERROR"
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    mstrLastError = "License validation error: " & Err.Description
    mbLicenseValid = False
    ValidateLicense = False
    
    If Not ErrorHandler Is Nothing Then
        ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "LicenseManager.ValidateLicense"
    End If
End Function

' Get license information as formatted string
Public Function GetLicenseInfo() As String
    On Error GoTo ErrorHandler
    
    GetLicenseInfo = ""
    
    If moLicenseValidator Is Nothing Then
        If Not CreateValidatorInstance() Then
            GetLicenseInfo = "License validator not available"
            Exit Function
        End If
    End If
    
    GetLicenseInfo = moLicenseValidator.GetLicenseInfo()
    Exit Function
    
ErrorHandler:
    GetLicenseInfo = "Error retrieving license info: " & Err.Description
    
    If Not ErrorHandler Is Nothing Then
        ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "LicenseManager.GetLicenseInfo"
    End If
End Function

' Get current Windows user
Public Function GetCurrentUser() As String
    On Error GoTo ErrorHandler
    
    GetCurrentUser = ""
    
    If moLicenseValidator Is Nothing Then
        If Not CreateValidatorInstance() Then
            GetCurrentUser = "Unknown"
            Exit Function
        End If
    End If
    
    GetCurrentUser = moLicenseValidator.GetCurrentUser()
    Exit Function
    
ErrorHandler:
    GetCurrentUser = "Error: " & Err.Description
    
    If Not ErrorHandler Is Nothing Then
        ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "LicenseManager.GetCurrentUser"
    End If
End Function

' Get number of authorized users
Public Function GetAuthorizedUserCount() As Integer
    On Error GoTo ErrorHandler
    
    GetAuthorizedUserCount = 0
    
    If moLicenseValidator Is Nothing Then
        If Not CreateValidatorInstance() Then
            Exit Function
        End If
    End If
    
    GetAuthorizedUserCount = moLicenseValidator.GetAuthorizedUserCount()
    Exit Function
    
ErrorHandler:
    GetAuthorizedUserCount = 0
    
    If Not ErrorHandler Is Nothing Then
        ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "LicenseManager.GetAuthorizedUserCount"
    End If
End Function

' Show license information dialog
Public Sub ShowLicenseDialog()
    On Error GoTo ErrorHandler
    
    Dim strMessage As String
    Dim strTitle As String
    
    strTitle = "ARES License Information"
    
    If mbLicenseValid Then
        strMessage = "? License Valid" & vbCrLf & vbCrLf
        strMessage = strMessage & GetLicenseInfo() & vbCrLf & vbCrLf
        strMessage = strMessage & "Current User: " & GetCurrentUser()
        
        MsgBox strMessage, vbInformation + vbOKOnly, strTitle
    Else
        strMessage = "? License Invalid" & vbCrLf & vbCrLf
        strMessage = strMessage & "Error: " & mstrLastError & vbCrLf & vbCrLf
        strMessage = strMessage & "Current User: " & GetCurrentUser() & vbCrLf & vbCrLf
        strMessage = strMessage & "Please contact your administrator to obtain a valid license."
        
        MsgBox strMessage, vbCritical + vbOKOnly, strTitle
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error displaying license information: " & Err.Description, vbCritical, strTitle
    
    If Not ErrorHandler Is Nothing Then
        ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "LicenseManager.ShowLicenseDialog"
    End If
End Sub

' Clean up COM object
Public Sub Cleanup()
    On Error Resume Next
    Set moLicenseValidator = Nothing
End Sub

' === PRIVATE METHODS ===

' Create instance of COM license validator
Private Function CreateValidatorInstance() As Boolean
    On Error GoTo ErrorHandler
    
    CreateValidatorInstance = False
    
    ' Try to create COM object
    Set moLicenseValidator = CreateObject("ARES.LicenseValidator")
    
    If moLicenseValidator Is Nothing Then
        mstrLastError = "Failed to create license validator COM object. Ensure AresLicenseValidator.dll is registered."
        Exit Function
    End If
    
    CreateValidatorInstance = True
    Exit Function
    
ErrorHandler:
    mstrLastError = "Error creating license validator: " & Err.Description & _
                    vbCrLf & "Ensure AresLicenseValidator.dll is properly registered with regasm."
    CreateValidatorInstance = False
    
    If Not ErrorHandler Is Nothing Then
        ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "LicenseManager.CreateValidatorInstance"
    End If
End Function


