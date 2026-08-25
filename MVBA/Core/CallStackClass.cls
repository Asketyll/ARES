' Class Module: CallStackClass
' Description: Lightweight manual call-stack tracer for ARES's key entry points. VBA has no native
'              call-stack introspection, so each instrumented procedure Push/Pop its own name onto
'              this fixed-size stack; ErrorHandlerClass reads it (read-only) to append the current
'              chain to every logged fault, and the LogCallStack key-in dumps LastSnapshot on demand.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: None
Option Explicit

' === CONSTANTS ===
' Bounds the stack against an undetected Push/Pop imbalance (e.g. a missed Pop on some future exit
' path). Comfortably above ElementChangeHandler.ProcessElement's own MAX_DEPTH=20 recursion cap.
Private Const MAX_STACK_DEPTH As Integer = 50

' === PRIVATE MEMBERS ===
Private msFrames(1 To MAX_STACK_DEPTH) As String
Private mnTop As Integer                ' 0 = empty; index of the current top frame otherwise
Private mbDepthCapped As Boolean        ' True once Push has been refused at least once since the stack last emptied
Private msLastSnapshot As String        ' Rendered chain at the last Push - see Push's comment

' === CLASS INITIALIZATION ===
Private Sub Class_Initialize()
    mnTop = 0
    mbDepthCapped = False
    msLastSnapshot = ""
End Sub

' === PUBLIC METHODS ===
' Push one frame (the calling procedure's own "Module.Proc" name), optionally tagged with the element
' it is operating on (ID/Type/GraphicGroup - what Asketyll used to dump by hand during debugging).
' Silently refuses past MAX_STACK_DEPTH rather than raising or logging - CallStackClass has no
' dependency on ErrorHandlerClass (the reverse dependency exists, for FormatErrorMessage to read this
' class), so it cannot log through it without creating a circular module dependency. A capped stack is
' still surfaced: the overflow is appended to the rendered Snapshot/LastSnapshot text instead.
Public Sub Push(ByVal ProcName As String, Optional ByVal oEl As element = Nothing)
    On Error Resume Next
    If mnTop >= MAX_STACK_DEPTH Then
        mbDepthCapped = True
        Exit Sub
    End If
    mnTop = mnTop + 1
    msFrames(mnTop) = ProcName & ElementContextSuffix(oEl)
    ' Re-rendered on every Push (not on Pop): calls nest before they unwind, so the last render
    ' before a top-level call fully returns is always its deepest/most-recent chain - exactly what
    ' the manual LogCallStack key-in is for (there is never a "live" stack to read at key-in time,
    ' since MicroStation is single-threaded/synchronous and no ARES procedure is still on the stack
    ' by the time the user regains the command line).
    msLastSnapshot = Snapshot()
End Sub

' "" when oEl is not supplied (default), so a bare Push "X" renders exactly as before. Otherwise
' "[ID=... Type=... GG=...]" - ID via DLongToString (never CStr - Element.ID is a DLong, see
' mvba-cheatsheet.md), Type/GraphicGroup as their raw Long values (no reverse enum-to-name helper exists
' in this project for MsdElementType). Formatted into a String immediately, never a stored element
' reference, so a later Snapshot/LastSnapshot read can never hit a stale/invalidated handle. On Error
' Resume Next: if any property read faults (e.g. an already-deleted element), the whole concatenation is
' abandoned and this returns "" rather than a truncated, misleading fragment.
Private Function ElementContextSuffix(ByVal oEl As element) As String
    On Error Resume Next
    If oEl Is Nothing Then Exit Function
    ElementContextSuffix = " [ID=" & DLongToString(oEl.ID) & " Type=" & CStr(oEl.Type) & " GG=" & CStr(oEl.GraphicGroup) & "]"
End Function

' Pop the most recently pushed frame. No-op on an already-empty stack (defensive; every call site
' pairs one Push with one Pop, but a fault mid-instrumentation must never make this go negative).
Public Sub Pop()
    On Error Resume Next
    If mnTop <= 0 Then Exit Sub
    msFrames(mnTop) = ""
    mnTop = mnTop - 1
    If mnTop = 0 Then mbDepthCapped = False
End Sub

' Current chain, chronological (outermost first), e.g. "A -> B -> C". "" when empty.
Public Function Snapshot() As String
    On Error Resume Next
    Dim i As Integer
    Dim sResult As String

    For i = 1 To mnTop
        If i > 1 Then sResult = sResult & " -> "
        sResult = sResult & msFrames(i)
    Next i

    If mbDepthCapped Then sResult = sResult & " -> [...MAX_STACK_DEPTH reached]"
    Snapshot = sResult
End Function

' The deepest/most recent chain seen since the stack last emptied - see Push's comment. What the
' LogCallStack key-in dumps: there is no way to read a "live" stack from a manual command in a
' single-threaded, synchronous host, since no ARES procedure is on the stack once the user has
' control back to type a key-in. This instead answers "what did ARES just run through".
Public Property Get LastSnapshot() As String
    LastSnapshot = msLastSnapshot
End Property

' Number of frames currently pushed (0 = empty). Test/diagnostic seam.
Public Property Get CurrentDepth() As Integer
    CurrentDepth = mnTop
End Property
