' Class Module: ReRegisterIdleHandler
' Description: Lightweight idle handler that re-registers IChangeTrackEvents after bulk operations.
'              Used in conjunction with ElementChangeHandler's bulk operation detection.
'              When a merge or reprojection is detected, ElementChangeHandler suspends tracking
'              and registers this handler. When MicroStation becomes idle (operation complete),
'              this handler re-enables change tracking automatically.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: BootLoader (ResumeChangeTracking)

Option Explicit

Implements IEnterIdleEvent

' Event handler for when MicroStation enters idle state
' This fires after the bulk operation completes and MicroStation is ready for user input
Private Sub IEnterIdleEvent_EnterIdle(ByVal Reserved As Long)
    On Error GoTo ErrorHandler

    ' Remove this handler first (one-shot handler)
    RemoveEnterIdleEventHandler Me

    ' Release the member reference held in ElementChangeHandler
    If Not ChangeHandler Is Nothing Then
        ChangeHandler.ClearReRegHandler
    End If

    ' Re-register change tracking via BootLoader
    ResumeChangeTracking

    ShowStatus "ARES: Change tracking resumed after bulk operation"
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "ReRegisterIdleHandler.IEnterIdleEvent_EnterIdle"
End Sub
