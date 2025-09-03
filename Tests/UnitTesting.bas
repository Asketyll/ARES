' Module: UnitTesting
' Description: Unit tests for ARES application with MsgBox output
' License: This project is licensed under the AGPL-3.0.

Option Explicit

Sub MasterTest()
    ConfigTest
    LangManagerTest
    UUIDTest
    ARES_VARTest
    CustomPropertyHandlerTest
    ErrorHandlerTest
End Sub

' Test 1: Configuration module tests
Sub ConfigTest()
    Dim Results As String
    Dim TestPassed As Boolean
    
    Results = "=== CONFIG TEST RESULTS ===" & vbCrLf & vbCrLf
    
    ' Test SetVar
    TestPassed = Config.SetVar("ARES_Unit_testing", "I'm a test unit variable")
    Results = Results & "• SetVar: " & IIf(TestPassed, " PASSED", " FAILED") & vbCrLf
    
    ' Test GetVar
    Dim Value As String
    Value = Config.GetVar("ARES_Unit_testing")
    TestPassed = (Value = "I'm a test unit variable")
    Results = Results & "• GetVar: " & IIf(TestPassed, " PASSED", " FAILED") & vbCrLf
    Results = Results & "  Retrieved: " & Value & vbCrLf
    
    ' Test RemoveValue
    TestPassed = Config.RemoveValue("ARES_Unit_testing")
    Results = Results & "• RemoveValue: " & IIf(TestPassed, " PASSED", " FAILED") & vbCrLf
    
    ' Verify removal
    Value = Config.GetVar("ARES_Unit_testing")
    TestPassed = (Value = "" Or Value = ARES_NAVD)
    Results = Results & "• Verify Removal: " & IIf(TestPassed, " PASSED", " FAILED")
    
    MsgBox Results, vbOKOnly + vbInformation, "Config Test Results"
End Sub

' Test 2: Language Manager tests
Sub LangManagerTest()
    Dim Results As String
    Dim Translation As String
    
    Results = "=== LANGUAGE MANAGER TEST RESULTS ===" & vbCrLf & vbCrLf
    
    ' Set language to English
    LangManager.InitializeTranslations
    
    ' Test with parameters
    Translation = LangManager.GetTranslation("VarRemoveConfirm", "ARES_Test ")
    Results = Results & "• Translation with Parameters:" & vbCrLf
    Results = Results & "  " & Translation & vbCrLf
    Results = Results & "  " & IIf(InStr(Translation, "ARES_Test") > 0, " PASSED", " FAILED")
    
    MsgBox Results, vbOKOnly + vbInformation, "Language Manager Test Results"
End Sub

' Test 3: UUID Generator tests
Sub UUIDTest()
    Dim Results As String
    Dim UUID1 As String, UUID2 As String
    
    Results = "=== UUID GENERATOR TEST RESULTS ===" & vbCrLf & vbCrLf
    
    ' Generate first UUID
    UUID1 = uuid.GenerateV1
    Results = Results & "• UUID Generation:" & vbCrLf
    Results = Results & "  UUID1: " & UUID1 & vbCrLf
    Results = Results & "  " & IIf(Len(UUID1) > 0, " PASSED", " FAILED") & vbCrLf & vbCrLf
    
    ' Generate second UUID (should be different)
    UUID2 = uuid.GenerateV1
    Results = Results & "• UUID Uniqueness:" & vbCrLf
    Results = Results & "  UUID2: " & UUID2 & vbCrLf
    Results = Results & "  " & IIf(UUID1 <> UUID2, " PASSED (Unique)", " FAILED (Duplicate)") & vbCrLf & vbCrLf
    
    ' Validate UUID format (8-4-4-4-12)
    Dim ValidFormat As Boolean
    ValidFormat = (Len(UUID1) = 36) And _
                  (Mid(UUID1, 9, 1) = "-") And _
                  (Mid(UUID1, 14, 1) = "-") And _
                  (Mid(UUID1, 19, 1) = "-") And _
                  (Mid(UUID1, 24, 1) = "-")
    Results = Results & "• UUID Format Validation:" & vbCrLf
    Results = Results & "  " & IIf(ValidFormat, " PASSED (Valid format)", " FAILED (Invalid format)")
    
    MsgBox Results, vbOKOnly + vbInformation, "UUID Test Results"
End Sub

' Test 4: ARES Variables tests
Sub ARES_VARTest()
    Dim Results As String
    Dim TestResult As Boolean
    Dim ARESConfigTest As New ARESConfigClass
    
    Results = "=== ARES VARIABLES TEST RESULTS ===" & vbCrLf & vbCrLf
    
    ' Initialize MS Variables
    TestResult = ARESConfigTest.Initialize
    Results = Results & "• Initialize ARESConfig: " & IIf(TestResult, " PASSED", " FAILED") & vbCrLf & vbCrLf

    ' Test Reset Variable
    TestResult = ARESConfigTest.ResetConfigVar("ARES_Round")
    Results = Results & "• Reset ARES_Round to default: " & IIf(TestResult, " PASSED", " FAILED") & vbCrLf
    Results = Results & "  Current value: " & ARESConfigTest.ARES_ROUNDS.Value & vbCrLf & vbCrLf
    
    ' Test Get Config Variable
    Dim configVar As ARES_MS_VAR_Class
    Set configVar = ARESConfigTest.GetConfigVar("ARES_Round")
    TestResult = Not (configVar Is Nothing)
    Results = Results & "• Get Config Variable: " & IIf(TestResult, " PASSED", " FAILED") & vbCrLf
    If TestResult Then
        Results = Results & "  Key: " & configVar.key & vbCrLf
        Results = Results & "  Default: " & configVar.defaultValue & vbCrLf
        Results = Results & "  Current: " & configVar.Value
    End If
    
    MsgBox Results, vbOKOnly + vbInformation, "ARES Variables Test Results"
End Sub

' Test 5: Custom Property Handler tests
Sub CustomPropertyHandlerTest()
    Dim Results As String
    Dim TestLibrary As ItemTypeLibrary
    Dim TestResult As Boolean
    
    Results = "=== CUSTOM PROPERTY HANDLER TEST RESULTS ===" & vbCrLf & vbCrLf
    
    On Error Resume Next
    
    ' Test Get/Create ItemTypeLibrary
    Set TestLibrary = CustomPropertyHandler.GetItemTypeLibrary()
    TestResult = Not (TestLibrary Is Nothing)
    Results = Results & "• Get/Create ItemTypeLibrary: " & IIf(TestResult, " PASSED", " FAILED") & vbCrLf
    
    If TestResult Then
        Results = Results & "  Library Name: " & TestLibrary.Name & vbCrLf & vbCrLf
    Else
        Results = Results & "  Error: " & Err.Description & vbCrLf & vbCrLf
    End If
    
    ' Test Delete ItemTypeLibrary
    TestResult = CustomPropertyHandler.DeleteItemTypeLibrary()
    Results = Results & "• Delete ItemTypeLibrary: " & IIf(TestResult, " PASSED", " FAILED") & vbCrLf
    
    ' Recreate for future tests
    Set TestLibrary = CustomPropertyHandler.GetItemTypeLibrary()
    TestResult = Not (TestLibrary Is Nothing)
    Results = Results & "• Recreate ItemTypeLibrary: " & IIf(TestResult, " PASSED", " FAILED")
    
    On Error GoTo 0
    
    MsgBox Results, vbOKOnly + vbInformation, "Custom Property Handler Test Results"
End Sub

' Test 6: Error Handler tests
Sub ErrorHandlerTest()
    Dim TestErrorHandler As New ErrorHandlerClass
    Dim Results As String
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    Dim ResultTxt As String
    Dim FileExists As Boolean
    
    Results = "=== ERROR HANDLER TEST RESULTS ===" & vbCrLf & vbCrLf

    ' Test 1: error logging
    TotalTests = TotalTests + 1
    If TestErrorHandler.HandleError("Test error message", 1001, "TestFunction", "UnitTesting") Then
        TestsPassed = TestsPassed + 1
        Results = Results & "• Test 1 - Basic error logging: PASSED" & vbCrLf
    Else
        Results = Results & "• Test 1 - Basic error logging: FAILED" & vbCrLf
    End If

    ' Test 2: GetLastLogEntry
    TotalTests = TotalTests + 1
    ResultTxt = StrReverse(TestErrorHandler.GetLastLogEntry)
    If Left(ResultTxt, 60) = StrReverse(" [UnitTesting] Error 1001 (TestFunction): Test error message") Then
        TestsPassed = TestsPassed + 1
        Results = Results & "• Test 2 - Get last log entry: PASSED" & vbCrLf
    Else
        Results = Results & "• Test 2 - Get last log entry: FAILED" & vbCrLf
    End If

    ' Test 3: remove error file
    TotalTests = TotalTests + 1
    FileExists = (Len(Dir(TestErrorHandler.LogFilePath)) > 0)
    If Not FileExists Then
        Results = Results & "• Test log file not fund !"
    Else
        Kill TestErrorHandler.LogFilePath
        FileExists = (Len(Dir(TestErrorHandler.LogFilePath)) > 0)
        If Not FileExists Then
            TestsPassed = TestsPassed + 1
            Results = Results & "• Test log file removed !"
        Else
            Results = Results & "• Test log file not removed !"
        End If
    End If

    ' Summary
    Results = Results & vbCrLf & "=== SUMMARY ===" & vbCrLf
    Results = Results & "Tests Passed: " & TestsPassed & "/" & TotalTests & vbCrLf
    Results = Results & "Success Rate: " & Format(TestsPassed / TotalTests * 100, "0.0") & " %"
    
    MsgBox Results, vbOKOnly + vbInformation, "Error Handler Test Results"
End Sub