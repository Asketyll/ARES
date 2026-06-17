' Module: BootLoader
' Description: Initializes the VBA project on load and manages global objects.
'              Also provides change tracking suspension for bulk operations.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: DGNOpenClose, ElementChangeHandler, LangManager, ErrorHandlerClass,
'               ElementInProcesseClass, ARESConfigClass, LicenseManager

Option Explicit

' === WIN32 API ===
' On VBA7 we use GetTickCount64 (returns unsigned 64-bit ms since system start, no practical wraparound).
' On legacy VBA6 we fall back to GetTickCount (signed 32-bit Long in VBA: wraps at ~24.85 days).
' Used for the license re-check throttle to avoid VBA Timer's midnight rollover (Issue #3).
#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongLong
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

' === LICENSE RE-CHECK CONSTANTS ===
Private Const LICENSE_RECHECK_INTERVAL_DEFAULT As Long = 3600
Private Const LICENSE_RECHECK_INTERVAL_MIN     As Long = 60
Private Const LICENSE_RECHECK_INTERVAL_MAX     As Long = 86400

' === GLOBAL OBJECT INSTANCES ===
Public ChangeHandler As ElementChangeHandler
Public ErrorHandler As ErrorHandlerClass
Public ElementInProcesse As New ElementInProcesseClass
Public ARESConfig As New ARESConfigClass

' === PRIVATE OBJECTS ===
Private moOpenClose As DGNOpenClose
Private mbLicenseChecked As Boolean
Private mbLicenseValid As Boolean
Private mbChangeTrackingSuspended As Boolean
Private mbChangeTrackingAttached As Boolean      ' Real attachment state of the change-track handler in MicroStation's list (decoupled from the bulk "suspended" flag)
Private mbIdleProcessingActive As Boolean
Private mbDGNHandlersInitialized As Boolean      ' True once InitializeDGNHandlers has run

' === LICENSE RE-CHECK STATE ===
#If VBA7 Then
Private mllLastLicenseCheckTicks As LongLong     ' GetTickCount64 value at last check (VBA7)
#Else
Private mlLastLicenseCheckTicks As Long          ' GetTickCount value at last check (VBA6, signed 32-bit, wraps ~24.85 days)
#End If
Private mlCachedIntervalMs As Long               ' Cached validated interval in ms (0 = not yet read)
Private mbLicenseInvalidatedNotified As Boolean  ' MsgBox shown once per True->False transition
Private mbLicenseChangeTrackingPaused As Boolean ' Set when handler removed due to invalid license

' Entry point when the project is loaded
' Initializes all global objects and event handlers required for ARES operation
Public Sub OnProjectLoad()
    On Error GoTo ErrorHandler

    ' Initialize the global error handler first (critical for other components)
    If Not InitializeErrorHandler() Then Exit Sub

    ' Validate license before initializing components
    If Not ValidateLicenseOnLoad() Then
        ShowLicenseFailureMessage
        ' Even on initial license failure, register the idle handler so periodic
        ' re-validation can still run and recover the session if the license becomes valid.
        ' DGN-event handlers are NOT wired here (mbDGNHandlersInitialized stays False);
        ' OnLicenseRecovered will wire them on the False -> True transition.
        InitializeInitialIdleHandler
        Exit Sub
    End If

    ' Initialize core components in dependency order
    If Not InitializeDGNHandlers() Then Exit Sub
    If Not InitializeInitialIdleHandler() Then Exit Sub

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
    
    ' Start the periodic re-check throttle from boot, regardless of outcome.
    ' This baseline is required so RecheckLicenseIfDue can measure elapsed time and trigger
    ' the False -> True recovery transition when the license becomes valid mid-session.
#If VBA7 Then
    mllLastLicenseCheckTicks = GetTickCount64
#Else
    mlLastLicenseCheckTicks = GetTickCount
#End If

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
    mbDGNHandlersInitialized = True
    InitializeDGNHandlers = True
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "BootLoader.InitializeDGNHandlers"
    InitializeDGNHandlers = False
End Function

' Public accessor used by recovery path to know whether DGN handlers were ever wired.
Public Function AreDGNHandlersInitialized() As Boolean
    AreDGNHandlersInitialized = mbDGNHandlersInitialized
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

    ' --- Step 1: reset all scalar flags first (cheap, cannot raise; keeps ErrorHandler available) ---
    mbChangeTrackingSuspended = False
    mbChangeTrackingAttached = False
    mbIdleProcessingActive = False
    mbDGNHandlersInitialized = False
    mbLicenseChecked = False
    mbLicenseValid = False
    mbLicenseInvalidatedNotified = False
    mbLicenseChangeTrackingPaused = False
    mlCachedIntervalMs = 0
#If VBA7 Then
    mllLastLicenseCheckTicks = 0
#Else
    mlLastLicenseCheckTicks = 0
#End If

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

' ========================================
' PERIODIC LICENSE RE-VALIDATION (Issue #2 mitigation, Story 1-2)
' ========================================

' Read and validate the configured re-check interval (in seconds).
' Falls back to default and logs WARNING if value is missing, non-numeric, or out of range.
Public Function GetLicenseRecheckIntervalSeconds() As Long
    On Error GoTo ErrorHandler

    Dim lValue As Long
    Dim strRaw As String

    GetLicenseRecheckIntervalSeconds = LICENSE_RECHECK_INTERVAL_DEFAULT

    If Not ARESConfig.IsInitialized Then Exit Function
    If ARESConfig.ARES_LICENSE_RECHECK_INTERVAL Is Nothing Then Exit Function

    strRaw = ARESConfig.ARES_LICENSE_RECHECK_INTERVAL.Value

    On Error Resume Next
    lValue = CLng(strRaw)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo ErrorHandler
        ErrorHandler.HandleError "Invalid ARES_License_Recheck_Interval value '" & strRaw & "', falling back to default " & LICENSE_RECHECK_INTERVAL_DEFAULT, _
                                 0, "BootLoader.GetLicenseRecheckIntervalSeconds", "WARNING"
        Exit Function
    End If
    On Error GoTo ErrorHandler

    If lValue < LICENSE_RECHECK_INTERVAL_MIN Or lValue > LICENSE_RECHECK_INTERVAL_MAX Then
        ErrorHandler.HandleError "ARES_License_Recheck_Interval=" & lValue & " out of range [" & LICENSE_RECHECK_INTERVAL_MIN & ".." & LICENSE_RECHECK_INTERVAL_MAX & "], falling back to default " & LICENSE_RECHECK_INTERVAL_DEFAULT, _
                                 0, "BootLoader.GetLicenseRecheckIntervalSeconds", "WARNING"
        Exit Function
    End If

    GetLicenseRecheckIntervalSeconds = lValue
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "BootLoader.GetLicenseRecheckIntervalSeconds"
    GetLicenseRecheckIntervalSeconds = LICENSE_RECHECK_INTERVAL_DEFAULT
End Function

' Re-validate the license if the throttle interval has elapsed since the last check.
' Returns True if a check was actually performed (regardless of outcome), False if throttled.
'
' The validated interval is cached in mlCachedIntervalMs to avoid re-reading and re-validating
' ARES_License_Recheck_Interval on every idle event (which would flood ErrorHandler with WARNINGs
' when the value is malformed). Cache is refreshed only when not yet initialized; callers can
' force a refresh by setting mlCachedIntervalMs = 0 (e.g. from a config-changed signal).
Public Function RecheckLicenseIfDue() As Boolean
    On Error GoTo ErrorHandler

    RecheckLicenseIfDue = False

    ' First-pass guard: if the boot baseline has not been established (e.g. license was
    ' invalid at boot), we cannot meaningfully measure elapsed time. Skip silently and
    ' let the recovery path elsewhere handle it.
#If VBA7 Then
    If mllLastLicenseCheckTicks = 0 Then Exit Function
#Else
    If mlLastLicenseCheckTicks = 0 Then Exit Function
#End If

    ' Lazy-load the validated interval (single read, single WARNING on misconfig).
    If mlCachedIntervalMs = 0 Then
        mlCachedIntervalMs = GetLicenseRecheckIntervalSeconds() * 1000
    End If

#If VBA7 Then
    Dim llNow As LongLong
    Dim llDelta As LongLong
    llNow = GetTickCount64
    llDelta = llNow - mllLastLicenseCheckTicks  ' Unsigned 64-bit; wraparound is not a practical concern.
    If llDelta >= CLngLng(mlCachedIntervalMs) Then
        PerformLicenseRecheck
        RecheckLicenseIfDue = True
    End If
#Else
    Dim lNow As Long
    Dim lDelta As Long
    lNow = GetTickCount
    lDelta = lNow - mlLastLicenseCheckTicks
    ' Signed 32-bit Long wraps at ~24.85 days; treat negative delta as due to force a single
    ' re-check then refresh the baseline. Documented limitation on legacy VBA6 hosts.
    If lDelta < 0 Or lDelta >= mlCachedIntervalMs Then
        PerformLicenseRecheck
        RecheckLicenseIfDue = True
    End If
#End If

    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "BootLoader.RecheckLicenseIfDue"
    RecheckLicenseIfDue = False
End Function

' Perform the actual license re-check and route transitions.
' Preserves previous mbLicenseValid state if validation throws.
Private Sub PerformLicenseRecheck()
    On Error GoTo ErrorHandler

    Dim bWasValid As Boolean
    Dim bNowValid As Boolean

    bWasValid = mbLicenseValid

    bNowValid = LicenseManager.ValidateLicense()
    mbLicenseValid = bNowValid
#If VBA7 Then
    mllLastLicenseCheckTicks = GetTickCount64
#Else
    mlLastLicenseCheckTicks = GetTickCount
#End If

    If bWasValid And Not bNowValid Then
        OnLicenseInvalidated
    ElseIf Not bWasValid And bNowValid Then
        OnLicenseRecovered
    End If

    Exit Sub

ErrorHandler:
    ' Preserve previous state on error; just log and continue.
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "BootLoader.PerformLicenseRecheck"
#If VBA7 Then
    mllLastLicenseCheckTicks = GetTickCount64
#Else
    mlLastLicenseCheckTicks = GetTickCount
#End If
End Sub

' Handle True -> False license transition: drain queue, remove change-tracking handler, notify user.
Private Sub OnLicenseInvalidated()
    On Error GoTo ErrorHandler

    ' Drain pending queue: do not process elements with an invalid license.
    If Not ElementInProcesse Is Nothing Then
        ElementInProcesse.Clear
    End If

    ' Remove change-tracking handler (only if not already suspended for bulk operations).
    If Not ChangeHandler Is Nothing And Not mbChangeTrackingSuspended Then
        DetachChangeTracking
        mbLicenseChangeTrackingPaused = True
    End If

    ' Show MsgBox once per transition.
    If Not mbLicenseInvalidatedNotified Then
        Dim strBody As String
        Dim strTitle As String
        If LangManager.IsInit Then
            strBody = GetTranslation("LicenseInvalidatedMidSession")
            strTitle = GetTranslation("LicenseRecheckTitle")
        Else
            strBody = "ARES license has become invalid during this session." & vbCrLf & _
                      "Features have been disabled until the license is restored." & vbCrLf & vbCrLf & _
                      "Reason: " & LicenseManager.LastError
            strTitle = "ARES - License Invalid"
        End If
        MsgBox strBody, vbCritical + vbOKOnly, strTitle
        mbLicenseInvalidatedNotified = True
    End If

    ShowStatus "ARES: License invalid - features disabled. Reason: " & LicenseManager.LastError
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "BootLoader.OnLicenseInvalidated"
End Sub

' Handle False -> True license transition: re-register change-tracking handler, notify user.
'
' Two recovery paths are supported:
'   (a) Mid-session invalidation -> recovery: mbLicenseChangeTrackingPaused = True. Just re-attach.
'   (b) Boot-time invalid license -> recovery: DGN handlers were never initialized. Wire them now.
' In both cases we reset the bulk-detection state on ChangeHandler before re-attaching to avoid
' stale counters / mid-window state from a previous (possibly long-ago) suspension.
Private Sub OnLicenseRecovered()
    On Error GoTo ErrorHandler

    ' Reset the one-shot notification so a future invalidation will MsgBox again.
    mbLicenseInvalidatedNotified = False

    Dim bNeedsAttach As Boolean
    bNeedsAttach = mbLicenseChangeTrackingPaused Or Not mbDGNHandlersInitialized

    If bNeedsAttach Then
        ' Ensure the DGN-event chain is wired (boot-invalid recovery path).
        If Not mbDGNHandlersInitialized Then
            If Not InitializeDGNHandlers() Then
                ErrorHandler.HandleError "Failed to initialize DGN handlers on license recovery", _
                                         0, "BootLoader.OnLicenseRecovered", "ERROR"
                Exit Sub
            End If
        End If

        If Not ChangeHandler Is Nothing Then
            ' Stale instance kept alive across invalidation may carry corrupted bulk-detection
            ' counters (mlCallCount, mdEntryTime). Reset before re-attaching.
            ChangeHandler.ResetBulkDetectionState
        End If
        ' AttachChangeTracking creates ChangeHandler if needed and attaches idempotently.
        AttachChangeTracking
        mbLicenseChangeTrackingPaused = False
    End If

    ShowStatus "ARES: License re-validated - features re-enabled"
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "BootLoader.OnLicenseRecovered"
End Sub