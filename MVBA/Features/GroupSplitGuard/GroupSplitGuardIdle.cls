' Class Module: GroupSplitGuardIdle
' Description: One-shot idle handler that asks GroupSplitGuard's pending question once MicroStation is idle -
'              never inside the change-track callback that detected the cut.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: GroupSplitGuard, ErrorHandlerClass (global ErrorHandler)

Option Explicit

Implements IEnterIdleEvent

Private Sub IEnterIdleEvent_EnterIdle(ByVal Reserved As Long)
    On Error GoTo ErrorHandler

    ' Off the idle list before the question: a modal dialog must not let this handler fire a second time.
    RemoveEnterIdleEventHandler Me
    GroupSplitGuard.ReleaseIdleHandler
    GroupSplitGuard.AskPending
    Exit Sub

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "GroupSplitGuardIdle.IEnterIdleEvent_EnterIdle"
End Sub
