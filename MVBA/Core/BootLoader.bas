' Module: BootLoader
' Description: Initializes the VBA project on load and manages global objects.
'              Also provides change tracking suspension for bulk operations.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: DGNOpenClose, ElementChangeHandler, LangManager, ErrorHandlerClass,
'               ElementInProcesseClass, ARESConfigClass

Option Explicit

' === GLOBAL OBJECT INSTANCES ===
Public ChangeHandler As ElementChangeHandler
Public ErrorHandler As ErrorHandlerClass
Public ElementInProcesse As New ElementInProcesseClass
Public ARESConfig As New ARESConfigClass

' === PRIVATE OBJECTS ===
Private moOpenClose As DGNOpenClose
Private mbChangeTrackingSuspended As Boolean
Private mbChangeTrackingAttached As Boolean      ' Real attachment state of the change-track handler in MicroStation's list (decoupled from the bulk "suspended" flag)
Private mbIdleProcessingActive As Boolean

' Entry point when the project is loaded
' Initializes all global objects and event handlers required for ARES operation
Public Sub OnProjectLoad()
    On Error GoTo ErrorHandler

    ' Initialize the global error handler first (critical for other components)
    If Not InitializeErrorHandler() Then Exit Sub

    ' Initialize core components in dependency order
    If Not InitializeDGNHandlers() Then Exit Sub
    If Not InitializeInitialIdleHandler() Then Exit Sub

    Exit Sub

ErrorHandler:
    ' Notify user about failure with detailed error information
    Dim sErrorMsg As String
    sErrorMsg = "Critical error during ARES initialization: " & Err.Description & vbCrLf & _
                  "Error Number: " & Err.Number & vbCrLf & _
                  "Source: " & Err.Source

    If LangManager.IsInit Then
        sErrorMsg = GetTranslation("BootFail") & vbCrLf & sErrorMsg
    End If

    MsgBox sErrorMsg, vbCritical + vbOKOnly, "ARES Initialization Failed"
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

' Initialize the INITIAL idle event handler (for project initialization only)
' Note: IdleHandlers for element processing are created dynamically by ElementChangeHandler
Private Function InitializeInitialIdleHandler() As Boolean
    On Error GoTo ErrorHandler

    Dim oInitialIdleHandler As IdleEventHandler

    ' Create and register idle event handler for initial project setup
    ' This handler will:
    '   1. Set the application caption
    '   2. Initialize translations
    '   3. Initialize ARESConfig
    '   4. Remove itself after execution
    Set oInitialIdleHandler = New IdleEventHandler
    AddEnterIdleEventHandler oInitialIdleHandler

    InitializeInitialIdleHandler = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "BootLoader.InitializeInitialIdleHandler"
    InitializeInitialIdleHandler = False
End Function

' Clean up global objects when project is unloaded
Public Sub OnProjectUnload()
    On Error Resume Next

    ' --- Step 1: reset all scalar flags first (cheap, cannot raise; keeps ErrorHandler available) ---
    mbChangeTrackingSuspended = False
    mbChangeTrackingAttached = False
    mbIdleProcessingActive = False

    ' --- Step 2: tear down objects in dependency order ---
    ' moOpenClose (DGN events) -> ChangeHandler (depends on ElementInProcesse) ->
    ' ElementInProcesse -> ARESConfig -> ErrorHandler (last, so previous teardown can still log).
    Set moOpenClose = Nothing
    Set ChangeHandler = Nothing
    Set ElementInProcesse = Nothing
    Set ARESConfig = Nothing
    Set ErrorHandler = Nothing
End Sub

' ========================================
' CHANGE TRACKING SUSPENSION - For bulk operations
' ========================================

' Idempotently attach the change-tracking handler to MicroStation's change-track list.
' Creates ChangeHandler if it does not exist yet. Safe to call repeatedly: the actual
' AddChangeTrackEventsHandler call only happens when not already attached, so it can never
' double-register (which is what made a late/stale ReRegisterIdleHandler dangerous before).
Public Sub AttachChangeTracking()
    On Error GoTo ErrorHandler

    If ChangeHandler Is Nothing Then Set ChangeHandler = New ElementChangeHandler

    If Not mbChangeTrackingAttached Then
        AddChangeTrackEventsHandler ChangeHandler
        mbChangeTrackingAttached = True
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "BootLoader.AttachChangeTracking"
End Sub

' Idempotently detach the change-tracking handler from MicroStation's change-track list.
Public Sub DetachChangeTracking()
    On Error GoTo ErrorHandler

    If mbChangeTrackingAttached And Not ChangeHandler Is Nothing Then
        RemoveChangeTrackEventsHandler ChangeHandler
        mbChangeTrackingAttached = False
    End If
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "BootLoader.DetachChangeTracking"
End Sub

' Real attachment state - reflects actual Add/Remove calls, NOT the bulk "suspended" flag.
Public Function IsChangeTrackingAttached() As Boolean
    IsChangeTrackingAttached = mbChangeTrackingAttached
End Function

' A new design file is now open: cleanly detach any previous handler, drop the bulk-suspend
' state, then attach a fresh handler. Single entry point for DGNOpenClose so the attachment
' bookkeeping stays consistent regardless of whatever bulk state was left over from the
' previous file (e.g. a suspend that was in progress when the conversion closed the file).
Public Sub ReinitChangeTrackingForNewFile()
    On Error GoTo ErrorHandler

    DetachChangeTracking                 ' remove the previous instance from MS's list (no-op if already detached)
    mbChangeTrackingSuspended = False
    Set ChangeHandler = New ElementChangeHandler
    AttachChangeTracking
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "BootLoader.ReinitChangeTrackingForNewFile"
End Sub

' Suspend change tracking to improve performance during bulk operations
' Usage: Run keyin "vba run [ARES]BootLoader.SuspendChangeTracking"
' Then perform merge/reprojection, then call ResumeChangeTracking
Public Sub SuspendChangeTracking()
    On Error GoTo ErrorHandler

    If mbChangeTrackingSuspended Then
        ShowStatus "ARES: Change tracking already suspended"
        Exit Sub
    End If

    If Not ChangeHandler Is Nothing Then
        DetachChangeTracking
        mbChangeTrackingSuspended = True
        ShowStatus "ARES: Change tracking SUSPENDED - perform bulk operation then resume"
    Else
        ShowStatus "ARES: No change handler to suspend"
    End If
    Exit Sub

ErrorHandler:
    ShowStatus "ARES: Error suspending change tracking: " & Err.Description
End Sub

' Resume change tracking after bulk operations
' Usage: Run keyin "vba run [ARES]BootLoader.ResumeChangeTracking"
' Also called automatically by ReRegisterIdleHandler after auto-suspend
'
' Decision to re-attach is driven by the REAL attachment state (mbChangeTrackingAttached), not by
' the bulk "suspended" flag. A file close (ResetSuspensionState) can clear the suspended flag while
' the handler is still detached by a bulk suspend; gating the re-attach on the suspended flag was
' silently skipping the AddChangeTrackEventsHandler -> change tracking stayed dead after the bulk.
Public Sub ResumeChangeTracking()
    On Error GoTo ErrorHandler

    AttachChangeTracking
    mbChangeTrackingSuspended = False
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "BootLoader.ResumeChangeTracking"
End Sub

' Check if change tracking is currently suspended
Public Function IsChangeTrackingSuspended() As Boolean
    IsChangeTrackingSuspended = mbChangeTrackingSuspended
End Function

' Mark change tracking as suspended (called by ElementChangeHandler during auto-suspend)
Public Sub MarkChangeTrackingSuspended()
    mbChangeTrackingSuspended = True
End Sub

' Called by IdleEventHandler at the start/end of ProcessPendingElements to suppress
' bulk detection triggered by ARES's own element modifications during idle processing.
Public Sub SetIdleProcessingActive(ByVal bActive As Boolean)
    mbIdleProcessingActive = bActive
End Sub

Public Function IsIdleProcessingActive() As Boolean
    IsIdleProcessingActive = mbIdleProcessingActive
End Function

' Reset suspension state on file close so stale ReRegisterIdleHandler won't re-register on next open
Public Sub ResetSuspensionState()
    mbChangeTrackingSuspended = False
End Sub
