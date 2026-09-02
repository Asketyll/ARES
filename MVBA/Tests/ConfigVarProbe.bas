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
' RESULT on OpenCities Map PowerView 2023, 2026-09-02:
'   0. PROVEN - raw and expanded reads of MS_DGNLIBLIST returned the SAME flattened string. The
'      API never exposes the $(...) definition even though the .ucf still holds it, so any
'      read-modify-write starts from already-flattened text.
'   1. PROVEN - writing "$(ARES_PROBE_A);BBB" stored "AAA;BBB": expansion happens ON WRITE.
'      This alone kills the idea, whatever step 2 says: the stored value can never keep a
'      reference, only the flattened text of what it pointed at.
'   2. NOT CONCLUSIVE - writing "$(ARES_PROBE_A);CCC" onto ARES_PROBE_A left the variable
'      UNDEFINED, but the variable only existed in-session, so the self-reference had no
'      lower-precedence definition to resolve to. MS_DGNLIBLIST does have one. Re-test with
'      ProbeSelfRefPersistedSetup / ProbeSelfRefPersistedCheck below.
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

' ==================================================================================================
'  SELF-REFERENCE, IN REAL CONDITIONS (two phases, a restart in between)
' ==================================================================================================
' Step 2 above wrote over a variable that existed only in this session, so "$(VAR);x" had nothing
' underneath to resolve to. MS_DGNLIBLIST is not in that situation: it is defined in configuration
' files read at startup, and a VBA write lands ON TOP of that. These two subs reproduce THAT.
'
' Phase 1 persists ARES_PROBE_A to the user configuration file, then you restart MicroStation so the
' definition comes back from the file rather than from this session. Phase 2 writes the self-reference
' over it and reports what survived.
'
' WHAT IT COSTS: phase 1 writes one line to Personal.ucf, and RemoveConfigurationVariable does NOT
' delete it (documented: it never touches the user configuration file). Phase 2 prints the line to
' delete by hand. It is a throwaway ARES_PROBE_A - never do this dance on a variable the site owns.
'
' NOTE ON WHAT IT CAN PROVE: at best it shows the self-reference resolves. It cannot rescue the idea,
' because step 1 already showed the write expands - the site's $(...) would still end up frozen as
' literal paths, one precedence level higher than the definition they came from.

' Phase 1: persist the probe variable, then RESTART MicroStation before running phase 2.
Public Sub ProbeSelfRefPersistedSetup()
    On Error GoTo ErrorHandler

    Application.ActiveWorkspace.AddConfigurationVariable PROBE_A, "AAA", True

    Debug.Print String(80, "=")
    Debug.Print PROBE_A & " written to the user configuration file with value AAA."
    Debug.Print "NOW RESTART MicroStation, then run ProbeSelfRefPersistedCheck."
    Debug.Print String(80, "=")
    Exit Sub

ErrorHandler:
    Debug.Print "SETUP ERROR " & Err.Number & " - " & Err.Description
End Sub

' Phase 2: run AFTER restarting. Writes the self-reference over the file-defined variable.
Public Sub ProbeSelfRefPersistedCheck()
    On Error GoTo ErrorHandler

    Dim sBefore As String
    Dim sAfter  As String

    sBefore = SafeValue(PROBE_A, True)
    Debug.Print String(80, "=")
    Debug.Print "before write : " & sBefore
    If InStr(sBefore, "AAA") = 0 Then
        Debug.Print "-> not defined from the file: run ProbeSelfRefPersistedSetup and RESTART first."
        Debug.Print String(80, "=")
        Exit Sub
    End If

    Application.ActiveWorkspace.AddConfigurationVariable PROBE_A, "$(" & PROBE_A & ");CCC", True
    sAfter = SafeValue(PROBE_A, True)

    Debug.Print "after write  : " & sAfter
    If InStr(sAfter, "AAA") > 0 And InStr(sAfter, "CCC") > 0 Then
        Debug.Print "-> the self-reference RESOLVES against a file-defined variable."
        Debug.Print "   Still not usable: step 1 showed the write expands, so the stored value is the"
        Debug.Print "   flattened text, not the reference."
    Else
        Debug.Print "-> the self-reference does NOT resolve even here. The route is dead outright."
    End If
    Debug.Print "CLEAN UP BY HAND: delete the '" & PROBE_A & "' lines from Personal.ucf."
    Debug.Print String(80, "=")
    Exit Sub

ErrorHandler:
    Debug.Print "CHECK ERROR " & Err.Number & " - " & Err.Description
    Debug.Print "(an error here is itself an answer, and " & PROBE_A & " may now be undefined)"
    Debug.Print "CLEAN UP BY HAND: delete the '" & PROBE_A & "' lines from Personal.ucf."
End Sub
