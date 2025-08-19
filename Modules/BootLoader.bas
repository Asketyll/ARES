' Module: BootLoader
' Description: Initializes the VBA project on load and manages global objects
' License: This project is licensed under the AGPL-3.0.
' Dependencies: DGNOpenClose, ElementChangeHandler, LangManager, ErrorHandlerClass, ElementInProcesseClass, ARESConfigClass
Option Explicit

' === GLOBAL OBJECT INSTANCES ===
Public ChangeHandler As ElementChangeHandler
Public ErrorHandler As ErrorHandlerClass
Public ElementInProcesse As ElementInProcesseClass
Public ARESConfig As New ARESConfigClass

' === PRIVATE OBJECTS ===
Private moOpenClose As DGNOpenClose
Private moIdleEventHandler As IdleEventHandler

' Entry point when the project is loaded
' Initializes all global objects and event handlers required for ARES operation
Public Sub OnProjectLoad()
    On Error GoTo ErrorHandler
    
    ' Initialize the global error handler first (critical for other components)
    If Not InitializeErrorHandler() Then Exit Sub
    
    ' Initialize core components in dependency order
    If Not InitializeDGNHandlers() Then Exit Sub
    If Not InitializeEventHandlers() Then Exit Sub
    
    ' Log successful initialization
    ErrorHandler.HandleError "ARES VBA project loaded successfully", 0, "BootLoader.OnProjectLoad", "INFO"
    
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
End Sub
