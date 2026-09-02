' Module: ConfigVarProbe
' Description: Platform probe answering ONE question: can VBA add an entry to a chained MicroStation
'              configuration variable without destroying what is already there?
'
'              Background: appending to MS_DGNLIBLIST by reading it back and rewriting it flattened the
'              site's $(_USTN_...) references into absolute paths and wiped the ARES entry. The only
'              known safe append is a ">" line in a .cfg/.ucf. This probe tests whether a THIRD route
'              exists - writing the literal string "$(VAR);newentry", the API equivalent of "VAR > x".
'
'              Two things must hold for that route to be usable, and neither can be read off the docs:
'                1. AddConfigurationVariable must store the string LITERALLY, not expand it on write.
'                2. A self-reference must resolve to the definition it overrides, not to itself.
'
'              SAFETY: every write here passes False as the third argument, so nothing reaches the user
'              configuration file and everything is gone when MicroStation closes. The probe variables
'              are removed at the end, on the error path too. MS_DGNLIBLIST is only ever READ.
'
' Usage: in the Immediate window, run ProbeConfigVarChaining and read the report there.
'
' RESULT on OpenCities Map PowerView 2023, 2026-09-02 - the route is DEAD, all three ways:
'   0. raw and expanded reads of MS_DGNLIBLIST returned the SAME flattened string; the API never
'      exposes the $(...) definition even though the .ucf still holds it.
'   1. writing "$(ARES_PROBE_A);BBB" stored "AAA;BBB" - expanded on write.
'   2. writing "$(ARES_PROBE_A);CCC" onto ARES_PROBE_A left the variable UNDEFINED. On
'      MS_DGNLIBLIST that would have deleted the site's whole DGNLib list.
' Kept in the repo as evidence, and to re-run on another MicroStation version before anyone
' reopens the question.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: none (Application.ActiveWorkspace only - deliberately no ARES module, so the probe
'               stays runnable on a station where ARES itself is misconfigured)
Option Explicit

Private Const PROBE_A As String = "ARES_PROBE_A"
Private Const PROBE_B As String = "ARES_PROBE_B"
Private Const REAL_VAR As String = "MS_DGNLIBLIST"   ' read-only here, never written

' Runs the probe and prints a verdict to the Immediate window.
Public Sub ProbeConfigVarChaining()
    On Error GoTo ErrorHandler

    Dim sRawB    As String
    Dim sExpB    As String
    Dim sSelf    As String
    Dim bLiteral As Boolean
    Dim bSelfRef As Boolean

    Debug.Print String(80, "=")
    Debug.Print "ARES config-variable probe - session only, nothing written to the .ucf"
    Debug.Print String(80, "=")

    ' --- 0. What the real chained variable looks like, raw vs expanded (READ ONLY) -----------------
    '     If the raw form still shows $(...), a raw read is possible and only the WRITE flattens.
    Debug.Print "0. " & REAL_VAR & " as MicroStation holds it"
    Debug.Print "   raw      : " & SafeValue(REAL_VAR, False)
    Debug.Print "   expanded : " & SafeValue(REAL_VAR, True)
    Debug.Print

    ' --- 1. Is a $(...) reference stored literally, or expanded on write? --------------------------
    Application.ActiveWorkspace.AddConfigurationVariable PROBE_A, "AAA", False
    Application.ActiveWorkspace.AddConfigurationVariable PROBE_B, "$(" & PROBE_A & ");BBB", False

    sRawB = SafeValue(PROBE_B, False)
    sExpB = SafeValue(PROBE_B, True)
    bLiteral = (InStr(sRawB, "$(") > 0)

    Debug.Print "1. writing """ & "$(" & PROBE_A & ");BBB"""
    Debug.Print "   raw      : " & sRawB
    Debug.Print "   expanded : " & sExpB
    Debug.Print "   -> stored " & IIf(bLiteral, "LITERALLY (good)", "ALREADY EXPANDED (route is dead)")
    Debug.Print

    ' --- 2. Does a self-reference resolve to the value being overridden? ---------------------------
    '     This is the whole point: "$(VAR);x" must mean "what VAR was, plus x".
    Application.ActiveWorkspace.AddConfigurationVariable PROBE_A, "$(" & PROBE_A & ");CCC", False

    sSelf = SafeValue(PROBE_A, True)
    bSelfRef = (InStr(sSelf, "AAA") > 0 And InStr(sSelf, "CCC") > 0)

    Debug.Print "2. rewriting " & PROBE_A & " as """ & "$(" & PROBE_A & ");CCC"""
    Debug.Print "   expanded : " & sSelf
    Debug.Print "   -> " & IIf(bSelfRef, "RESOLVES to the previous value (good)", "does NOT resolve (route is dead)")
    Debug.Print

    ' --- Verdict ----------------------------------------------------------------------------------
    Debug.Print String(80, "-")
    If bLiteral And bSelfRef Then
        Debug.Print "VERDICT: a non-destructive append from VBA IS possible."
        Debug.Print "         Write the literal ""$(" & REAL_VAR & ");C:/ARES/Rsc/*.dgnlib""."
        Debug.Print "         Still to weigh before using it: the value lands at User level, so confirm"
        Debug.Print "         on a real station that the chain resolves the same before and after."
    Else
        Debug.Print "VERDICT: no safe append from VBA. Keep the "">"" line in the .ucf as the only route,"
        Debug.Print "         and let ARES check and report rather than write."
    End If
    Debug.Print String(80, "-")

    Cleanup
    Exit Sub

ErrorHandler:
    Debug.Print "PROBE ERROR " & Err.Number & " - " & Err.Description
    Debug.Print "(an error on the self-reference step is itself an answer: the route is dead)"
    Cleanup
End Sub

' ConfigurationVariableValue raises when the variable is not defined, so every read goes through here.
Private Function SafeValue(ByVal sName As String, ByVal bExpand As Boolean) As String
    On Error GoTo NotDefined

    If Not Application.ActiveWorkspace.IsConfigurationVariableDefined(sName) Then
        SafeValue = "<undefined>"
        Exit Function
    End If

    SafeValue = Application.ActiveWorkspace.ConfigurationVariableValue(sName, bExpand)
    If Len(SafeValue) = 0 Then SafeValue = "<empty>"
    Exit Function

NotDefined:
    SafeValue = "<error " & Err.Number & ": " & Err.Description & ">"
End Function

' Drops both probe variables. They were session-only to begin with; this just keeps the session clean.
Private Sub Cleanup()
    On Error Resume Next
    Application.ActiveWorkspace.RemoveConfigurationVariable PROBE_A
    Application.ActiveWorkspace.RemoveConfigurationVariable PROBE_B
    On Error GoTo 0
    Debug.Print "probe variables removed."
End Sub
