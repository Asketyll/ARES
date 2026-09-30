' Module: Messaging
' Description: The one door ARES uses to post to MicroStation's Message Center. Fail-silent by design.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: none
Option Explicit

' Posts one entry to the Message Center. Never raises and never calls ErrorHandler.HandleError:
' HandleError posts through here, so reporting a failure of this helper would recurse, and a Message
' Center failure must never stop the caller (the .log write in particular).
Public Sub PostMessage(ByVal Message As String, _
                       Optional ByVal Details As String = "", _
                       Optional ByVal Priority As MsdMessageCenterPriority = msdMessageCenterPriorityInfo)
    On Error Resume Next
    MessageCenter.AddMessage Message, Details, Priority, False
End Sub
