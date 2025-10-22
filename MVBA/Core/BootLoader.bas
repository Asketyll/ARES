' Module: BootLoader
' Description: Initializes the VBA project on load and manages global objects
' License: This project is licensed under the AGPL-3.0.
' Dependencies: DGNOpenClose, ElementChangeHandler, LangManager, ErrorHandlerClass, ElementInProcesseClass, ARESConfigClass, LicenseManager
Option Explicit

' === GLOBAL OBJECT INSTANCES ===
Public ChangeHandler As ElementChangeHandler
Public ErrorHandler As ErrorHandlerClass
Public ElementInProcesse As New ElementInProcesseClass
Public ARESConfig As New ARESConfigClass

' === PRIVATE OBJECTS ===
Private moOpenClose As DGNOpenClose
Private moIdleEventHandler As IdleEventHandler
Private mbLicenseChecked As Boolean
Private mbLicenseValid As Boolean

' Entry point when the project is loaded
' Initializes all global objects and event handlers required for ARES operation
Public Sub OnProjectLoad()
    On Error GoTo ErrorHandler
    
    ' Initialize the global error handler first (critical for other components)
    If Not InitializeErrorHandler() Then Exit Sub

    ' Validate license before initializing components
    If Not ValidateLicenseOnLoad() Then
        ShowLicenseFailureMessage
        Exit Sub
    End If
    
    ' Initialize core components in dependency order
    If Not InitializeDGNHandlers() Then Exit Sub
    If Not InitializeEventHandlers() Then Exit Sub
    
    Exit Sub

ErrorHandler:
    ' Notify user about failure with detailed error information
    Dim strErrorMsg As String
    strErrorMsg = "Critical error during ARES initialization: " & Err.Description & vbCrLf & _
                  "Error Number: " & Err.Number & vbCrLf & _
                  "Source: " & Err.Source
    
    If LangManager.IsInit Then
        strErrorMsg = GetTranslation("BootFail") & vbCrLf & strErrorMsg
    End If
    
    MsgBox strErrorMsg, vbCritical + vbOKOnly, "ARES Initialization Failed"
End Sub

' Initialize the global error handler
Private Function InitializeErrorHandler() As Boolean
    On Error Resume Next
    
    Set ErrorHandler = New ErrorHandlerClass
    InitializeErrorHandler = (Err.Number = 0)
    
    If Err.Number <> 0 Then
        MsgBox "Failed to initialize ErrorHandler: " & Err.Description, vbCritical
    End If
End Function

' Validate license on application load
Private Function ValidateLicenseOnLoad() As Boolean
    On Error GoTo ErrorHandler
    
    ValidateLicenseOnLoad = False
    mbLicenseChecked = False
    mbLicenseValid = False
    
    ' Validate the license
    mbLicenseValid = LicenseManager.ValidateLicense()
    mbLicenseChecked = True
    
    If mbLicenseValid Then
        ValidateLicenseOnLoad = True
    Else
        ErrorHandler.HandleError "License validation failed: " & LicenseManager.LastError, 0, "BootLoader.ValidateLicenseOnLoad", "ERROR"
    End If
    
    Exit Function
    
ErrorHandler:
    mbLicenseChecked = True
    mbLicenseValid = False
    ValidateLicenseOnLoad = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "BootLoader.ValidateLicenseOnLoad"
End Function

' Show license failure message to user
Private Sub ShowLicenseFailureMessage()
    On Error Resume Next
    
    Dim strMessage As String
    Dim strTitle As String
    
    strTitle = "ARES - License Validation Failed"
    strMessage = "ARES cannot start because license validation failed." & vbCrLf & vbCrLf
    strMessage = strMessage & "Error: " & LicenseManager.LastError & vbCrLf & vbCrLf
    strMessage = strMessage & "Current User: " & LicenseManager.GetCurrentUser() & vbCrLf & vbCrLf
    strMessage = strMessage & "Possible causes:" & vbCrLf
    strMessage = strMessage & "• License file not found on network" & vbCrLf
    strMessage = strMessage & "• User not authorized in license" & vbCrLf
    strMessage = strMessage & "• Domain mismatch" & vbCrLf
    strMessage = strMessage & "• Invalid license signature" & vbCrLf & vbCrLf
    strMessage = strMessage & "Please contact your system administrator."
    
    MsgBox strMessage, vbCritical + vbOKOnly, strTitle
    ShowStatus "ARES disabled - License validation failed"
End Sub

' Initialize DGN file handlers
Private Function InitializeDGNHandlers() As Boolean
    On Error GoTo ErrorHandler
    
    Set moOpenClose = New DGNOpenClose
    InitializeDGNHandlers = True
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "BootLoader.InitializeDGNHandlers"
    InitializeDGNHandlers = False
End Function

' Initialize event handlers
Private Function InitializeEventHandlers() As Boolean
    On Error GoTo ErrorHandler
    
    ' Create and register idle event handler
    Set moIdleEventHandler = New IdleEventHandler
    AddEnterIdleEventHandler moIdleEventHandler
    
    InitializeEventHandlers = True
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "BootLoader.InitializeEventHandlers"
    InitializeEventHandlers = False
End Function

' Public function to check if license is valid (can be called from other modules)
Public Function IsLicenseValid() As Boolean
    IsLicenseValid = mbLicenseValid
End Function

' Public function to get license status info
Public Function GetLicenseStatus() As String
    On Error Resume Next
    
    If Not mbLicenseChecked Then
        GetLicenseStatus = "License not checked"
    ElseIf mbLicenseValid Then
        GetLicenseStatus = "Valid - " & LicenseManager.GetCurrentUser()
    Else
        GetLicenseStatus = "Invalid - " & LicenseManager.LastError
    End If
End Function

' Public sub to show license information dialog (can be called from command/macro)
Public Sub ShowLicenseInfo()
    On Error Resume Next
    
    If mbLicenseChecked Then
        LicenseManager.ShowLicenseDialog
    Else
        MsgBox "License has not been validated yet.", vbInformation, "ARES License"
    End If
End Sub

' Clean up global objects when project is unloaded
Public Sub OnProjectUnload()
    On Error Resume Next
    
    ' Clean up objects in reverse order of initialization
    Set moIdleEventHandler = Nothing
    Set moOpenClose = Nothing
    Set ElementInProcesse = Nothing
    Set ChangeHandler = Nothing
    Set ARESConfig = Nothing
    Set ErrorHandler = Nothing
    
    mbLicenseChecked = False
    mbLicenseValid = False
End Sub
