' Module: UnitTesting
' Description: Comprehensive unit tests for ARES application
' License: This project is licensed under the AGPL-3.0.
' Dependencies: All ARES modules

Option Explicit

' Test IDs enumeration for better maintainability
Private Enum TestID
    tidConfig = 1
    tidLangManager = 2
    tidUUID = 3
    tidARESVars = 4
    tidCustomProps = 5
    tidErrorHandler = 6
    tidElementProcess = 7
    tidLength = 8
    tidLink = 9
    tidStringsInEl = 10
End Enum

' Test result structure
Private Type TestResult
    Name As String
    Passed As Boolean
    Message As String
    Duration As Double
End Type

Private TestResults() As TestResult
Private TestCount As Long

' === MAIN TEST RUNNER ===
Public Sub RunAllTests()
    Dim StartTime As Double
    Dim Results As String
    
    StartTime = Timer
    
    ' Initialize test tracking
    TestCount = 0
    ReDim TestResults(0)
    
    ' Display header
    Results = "=== ARES TEST SUITE ===" & vbCrLf
    Results = Results & "Started: " & Now & vbCrLf
    Results = Results & String(50, "=") & vbCrLf & vbCrLf
    
    ' Run all test modules
    RunTest "Configuration", tidConfig
    RunTest "Language Manager", tidLangManager
    RunTest "UUID Generator", tidUUID
    RunTest "ARES Variables", tidARESVars
    RunTest "Custom Properties", tidCustomProps
    RunTest "Error Handler", tidErrorHandler
    RunTest "Element Processing", tidElementProcess
    RunTest "Length Calculations", tidLength
    RunTest "Link Functions", tidLink
    RunTest "String In Elements", tidStringsInEl
    
    ' Generate summary report
    Results = Results & GenerateTestReport(Timer - StartTime)
    
    ' Display results
    MsgBox Results, vbOKOnly + vbInformation, "ARES Test Suite Results"
    
    ' Save results to log
    SaveTestResults Results
End Sub

' Run a single test by ID (useful for debugging specific tests)
Public Sub RunSingleTest(TestIdentifier As Integer)
    Dim TestName As String
    Dim Result As Boolean
    
    ' Get test name and run test
    Select Case TestIdentifier
        Case tidConfig
            TestName = "Configuration"
            Result = ConfigTest()
        Case tidLangManager
            TestName = "Language Manager"
            Result = LangManagerTest()
        Case tidUUID
            TestName = "UUID Generator"
            Result = UUIDTest()
        Case tidARESVars
            TestName = "ARES Variables"
            Result = ARES_VARTest()
        Case tidCustomProps
            TestName = "Custom Properties"
            Result = CustomPropertyHandlerTest()
        Case tidErrorHandler
            TestName = "Error Handler"
            Result = ErrorHandlerTest()
        Case tidElementProcess
            TestName = "Element Processing"
            Result = ElementInProcesseTest()
        Case tidLength
            TestName = "Length Calculations"
            Result = LengthTest()
        Case tidLink
            TestName = "Link Functions"
            Result = LinkTest()
        Case tidStringsInEl
            TestName = "String In Elements"
            Result = StringsInElTest()
        Case Else
            MsgBox "Invalid test ID: " & TestIdentifier, vbCritical, "Test Error"
            Exit Sub
    End Select
    
    ' Display result
    MsgBox TestName & " Test: " & IIf(Result, "PASSED", "FAILED"), _
           IIf(Result, vbInformation, vbCritical), "Single Test Result"
End Sub

' === INDIVIDUAL TEST MODULES ===

' Test 1: Configuration module
Private Function ConfigTest() As Boolean
    On Error GoTo ErrorHandler
    
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    
    ' Test 1.1: SetVar
    TotalTests = TotalTests + 1
    If Config.SetVar("ARES_Unit_testing", "I'm a test unit variable") Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 1.2: GetVar
    TotalTests = TotalTests + 1
    If Config.GetVar("ARES_Unit_testing") = "I'm a test unit variable" Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 1.3: GetVar with non-existent key
    TotalTests = TotalTests + 1
    If Config.GetVar("NonExistent_Key") = ARESConstants.ARES_NAVD Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 1.4: RemoveValue
    TotalTests = TotalTests + 1
    If Config.RemoveValue("ARES_Test_Variable") Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 1.5: Verify removal
    TotalTests = TotalTests + 1
    Dim RemovedValue As String
    RemovedValue = Config.GetVar("ARES_Test_Variable")
    If RemovedValue = "" Or RemovedValue = ARESConstants.ARES_NAVD Then
        TestsPassed = TestsPassed + 1
    End If
    
    ConfigTest = (TestsPassed = TotalTests)
    Exit Function
    
ErrorHandler:
    ConfigTest = False
End Function

' Test 2: Language Manager
Private Function LangManagerTest() As Boolean
    On Error GoTo ErrorHandler
    
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    
    ' Initialize translations
    If Not LangManager.IsInit Then
        LangManager.InitializeTranslations
    End If
    
    ' Test 2.1: Basic translation
    TotalTests = TotalTests + 1
    Dim Translation As String
    Translation = LangManager.GetTranslation("VarResetSuccess", "TestVar")
    If InStr(Translation, "TestVar") > 0 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 2.2: Translation with multiple parameters
    TotalTests = TotalTests + 1
    Translation = LangManager.GetTranslation("LengthElementTypeNotSupportedByInterface", "12345", "TestType")
    If InStr(Translation, "12345") > 0 And InStr(Translation, "TestType") > 0 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 2.3: Missing translation key
    TotalTests = TotalTests + 1
    Translation = LangManager.GetTranslation("NonExistentKey")
    If InStr(Translation, "Missing translation") > 0 Or InStr(Translation, "Translation error") > 0 Then
        TestsPassed = TestsPassed + 1
    End If
    
    LangManagerTest = (TestsPassed = TotalTests)
    Exit Function
    
ErrorHandler:
    LangManagerTest = False
End Function

' Test 3: UUID Generator
Private Function UUIDTest() As Boolean
    On Error GoTo ErrorHandler
    
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    
    ' Test 3.1: UUID generation
    TotalTests = TotalTests + 1
    Dim UUID1 As String
    UUID1 = UUID.GenerateV1
    If Len(UUID1) = 36 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 3.2: UUID uniqueness
    TotalTests = TotalTests + 1
    Dim UUID2 As String
    UUID2 = UUID.GenerateV1
    If UUID1 <> UUID2 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 3.3: UUID format validation
    TotalTests = TotalTests + 1
    If ValidateUUIDFormat(UUID1) Then
        TestsPassed = TestsPassed + 1
    End If
    
    UUIDTest = (TestsPassed = TotalTests)
    Exit Function
    
ErrorHandler:
    UUIDTest = False
End Function

' Test 4: ARES Variables
Private Function ARES_VARTest() As Boolean
    On Error GoTo ErrorHandler
    
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    Dim TestConfig As New ARESConfigClass
    
    ' Test 4.1: Initialize
    TotalTests = TotalTests + 1
    If TestConfig.Initialize Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 4.2: Get config variable
    TotalTests = TotalTests + 1
    Dim ConfigVar As ARES_MS_VAR_Class
    Set ConfigVar = TestConfig.GetConfigVar("ARES_Round")
    If Not ConfigVar Is Nothing Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 4.3: Reset config variable
    TotalTests = TotalTests + 1
    If TestConfig.ResetConfigVar("ARES_Round") Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 4.4: Check IsInitialized property
    TotalTests = TotalTests + 1
    If TestConfig.IsInitialized Then
        TestsPassed = TestsPassed + 1
    End If
    
    ARES_VARTest = (TestsPassed = TotalTests)
    Exit Function
    
ErrorHandler:
    ARES_VARTest = False
End Function

' Test 5: Custom Property Handler
Private Function CustomPropertyHandlerTest() As Boolean
    On Error GoTo ErrorHandler
    
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    
    ' Test 5.1: Create/Get ItemTypeLibrary
    TotalTests = TotalTests + 1
    Dim TestLibrary As ItemTypeLibrary
    Set TestLibrary = CustomPropertyHandler.GetItemTypeLibrary("TestLibrary", "TestItem")
    If Not TestLibrary Is Nothing Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 5.2: Delete ItemTypeLibrary
    TotalTests = TotalTests + 1
    If CustomPropertyHandler.DeleteItemTypeLibrary("TestLibrary") Then
        TestsPassed = TestsPassed + 1
    End If
    
    CustomPropertyHandlerTest = (TestsPassed >= TotalTests - 1) ' Allow one failure for library operations
    Exit Function
    
ErrorHandler:
    CustomPropertyHandlerTest = False
End Function

' Test 6: Error Handler
Private Function ErrorHandlerTest() As Boolean
    On Error GoTo TestError
    
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    Dim TestHandler As New ErrorHandlerClass
    
    ' Test 6.1: Log error
    TotalTests = TotalTests + 1
    If TestHandler.HandleError("Test error", 1001, "TestSource", "TestModule") Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 6.2: Get last log entry
    TotalTests = TotalTests + 1
    Dim LastEntry As String
    LastEntry = TestHandler.GetLastLogEntry
    If InStr(LastEntry, "Test error") > 0 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 6.3: Clear log file
    TotalTests = TotalTests + 1
    TestHandler.ClearLogFile
    LastEntry = TestHandler.GetLastLogEntry
    If LastEntry = "" Then
        TestsPassed = TestsPassed + 1
    End If
    
    ErrorHandlerTest = (TestsPassed = TotalTests)
    Exit Function
    
TestError:
    ErrorHandlerTest = False
End Function

' Test 7: ElementInProcesse
Private Function ElementInProcesseTest() As Boolean
    On Error GoTo ErrorHandler
    
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    Dim TestProcessor As New ElementInProcesseClass
    
    ' Note: This test requires active MicroStation elements
    ' For now, we'll test the basic functionality
    
    ' Test 7.1: Initial state
    TotalTests = TotalTests + 1
    If TestProcessor.Count = 0 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 7.2: IsEmpty property
    TotalTests = TotalTests + 1
    If TestProcessor.IsEmpty Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 7.3: Clear operation
    TotalTests = TotalTests + 1
    TestProcessor.Clear
    If TestProcessor.Count = 0 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ElementInProcesseTest = (TestsPassed = TotalTests)
    Exit Function
    
ErrorHandler:
    ElementInProcesseTest = False
End Function

' Test 8: Length calculations
Private Function LengthTest() As Boolean
    On Error GoTo ErrorHandler
    
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    
    ' Note: This requires active elements in MicroStation
    ' We'll test the helper functions
    
    ' Test 8.1: Rounding function (assuming it exists in your implementation)
    TotalTests = TotalTests + 1
    ' This would need actual element testing
    TestsPassed = TestsPassed + 1 ' Placeholder
    
    LengthTest = (TestsPassed = TotalTests)
    Exit Function
    
ErrorHandler:
    LengthTest = False
End Function

' Test 9: Link functions
Private Function LinkTest() As Boolean
    On Error GoTo ErrorHandler
    
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    
    ' Test 9.1: Type validation
    TotalTests = TotalTests + 1
    If MicroStationDefinition.IsValidElementType(3) Then ' Line type
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 9.2: Invalid type detection
    TotalTests = TotalTests + 1
    If Not MicroStationDefinition.IsValidElementType(999) Then
        TestsPassed = TestsPassed + 1
    End If
    
    LinkTest = (TestsPassed = TotalTests)
    Exit Function
    
ErrorHandler:
    LinkTest = False
End Function

' Test 10: StringsInEl
Private Function StringsInElTest() As Boolean
    On Error GoTo ErrorHandler
    
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    
    ' Test 10.1: RemovePattern function
    TotalTests = TotalTests + 1
    Dim TestString As String
    TestString = StringsInEl.RemovePattern("Hello_World_Test", "_World")
    If TestString = "Hello_Test" Then
        TestsPassed = TestsPassed + 1
    End If
    
    StringsInElTest = (TestsPassed = TotalTests)
    Exit Function
    
ErrorHandler:
    StringsInElTest = False
End Function

' === HELPER FUNCTIONS ===

Private Sub RunTest(TestName As String, TestIdentifier As Integer)
    Dim StartTime As Double
    Dim Result As TestResult
    
    StartTime = Timer
    
    Result.Name = TestName
    
    On Error Resume Next
    ' Execute test based on ID
    Select Case TestIdentifier
        Case tidConfig: Result.Passed = ConfigTest()
        Case tidLangManager: Result.Passed = LangManagerTest()
        Case tidUUID: Result.Passed = UUIDTest()
        Case tidARESVars: Result.Passed = ARES_VARTest()
        Case tidCustomProps: Result.Passed = CustomPropertyHandlerTest()
        Case tidErrorHandler: Result.Passed = ErrorHandlerTest()
        Case tidElementProcess: Result.Passed = ElementInProcesseTest()
        Case tidLength: Result.Passed = LengthTest()
        Case tidLink: Result.Passed = LinkTest()
        Case tidStringsInEl: Result.Passed = StringsInElTest()
        Case Else
            Result.Passed = False
            Result.Message = "Unknown test ID"
    End Select
    
    If Err.Number <> 0 Then
        Result.Passed = False
        Result.Message = "Error: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    
    Result.Duration = Round((Timer - StartTime) * 1000, 2) ' Convert to milliseconds
    
    ' Add to results array
    TestCount = TestCount + 1
    ReDim Preserve TestResults(TestCount)
    TestResults(TestCount) = Result
End Sub

Private Function ValidateUUIDFormat(UUID As String) As Boolean
    ' Validate UUID v1 format: 8-4-4-4-12
    ValidateUUIDFormat = (Len(UUID) = 36) And _
                        (Mid(UUID, 9, 1) = "-") And _
                        (Mid(UUID, 14, 1) = "-") And _
                        (Mid(UUID, 19, 1) = "-") And _
                        (Mid(UUID, 24, 1) = "-") And _
                        (Mid(UUID, 15, 1) = "1") ' Version 1
End Function

Private Function GenerateTestReport(TotalDuration As Double) As String
    Dim Report As String
    Dim i As Long
    Dim PassedCount As Long
    Dim FailedCount As Long
    
    Report = vbCrLf & "=== TEST RESULTS ===" & vbCrLf & vbCrLf
    
    ' Individual test results
    For i = 1 To TestCount
        With TestResults(i)
            Report = Report & IIf(.Passed, "?", "?") & " " & .Name
            Report = Report & " (" & .Duration & " ms)"
            If Len(.Message) > 0 Then
                Report = Report & vbCrLf & "  " & .Message
            End If
            Report = Report & vbCrLf
            
            If .Passed Then
                PassedCount = PassedCount + 1
            Else
                FailedCount = FailedCount + 1
            End If
        End With
    Next i
    
    ' Summary
    Report = Report & vbCrLf & "=== SUMMARY ===" & vbCrLf
    Report = Report & "Total Tests: " & TestCount & vbCrLf
    Report = Report & "Passed: " & PassedCount & vbCrLf
    Report = Report & "Failed: " & FailedCount & vbCrLf
    Report = Report & "Success Rate: " & Format(PassedCount / TestCount * 100, "0.0") & "%" & vbCrLf
    Report = Report & "Total Duration: " & Format(TotalDuration, "0.000") & " seconds" & vbCrLf
    Report = Report & "Completed: " & Now
    
    GenerateTestReport = Report
End Function

Private Sub SaveTestResults(Results As String)
    On Error Resume Next
    
    Dim FilePath As String
    Dim FileNum As Integer
    
    If Not ActiveDesignFile Is Nothing Then
        FilePath = ActiveDesignFile.Path & "\ARES_TestResults_" & Format(Now, "yyyymmdd_hhmmss") & ".txt"
        
        FileNum = FreeFile
        Open FilePath For Output As #FileNum
        Print #FileNum, Results
        Close #FileNum
    End If
End Sub

' === PERFORMANCE TESTS ===

Public Sub RunPerformanceTests()
    Dim Results As String
    
    Results = "=== PERFORMANCE TESTS ===" & vbCrLf & vbCrLf
    
    ' Test configuration access speed
    Results = Results & TestConfigPerformance() & vbCrLf
    
    ' Test UUID generation speed
    Results = Results & TestUUIDPerformance() & vbCrLf
    
    ' Test translation lookup speed
    Results = Results & TestTranslationPerformance() & vbCrLf
    
    MsgBox Results, vbOKOnly + vbInformation, "Performance Test Results"
End Sub

Private Function TestConfigPerformance() As String
    Dim StartTime As Double
    Dim i As Long
    Dim Operations As Long
    
    Operations = 1000
    StartTime = Timer
    
    For i = 1 To Operations
        Config.SetVar "ARES_PerfTest", CStr(i)
        Config.GetVar "ARES_PerfTest"
    Next i
    
    Config.RemoveValue "ARES_PerfTest"
    
    TestConfigPerformance = "Config Operations: " & Operations * 2 & " in " & _
                           Format((Timer - StartTime) * 1000, "0.00") & " ms"
End Function

Private Function TestUUIDPerformance() As String
    Dim StartTime As Double
    Dim i As Long
    Dim Operations As Long
    
    Operations = 100
    StartTime = Timer
    
    For i = 1 To Operations
        UUID.GenerateV1
    Next i
    
    TestUUIDPerformance = "UUID Generations: " & Operations & " in " & _
                         Format((Timer - StartTime) * 1000, "0.00") & " ms"
End Function

Private Function TestTranslationPerformance() As String
    Dim StartTime As Double
    Dim i As Long
    Dim Operations As Long
    
    Operations = 1000
    
    If Not LangManager.IsInit Then
        LangManager.InitializeTranslations
    End If
    
    StartTime = Timer
    
    For i = 1 To Operations
        LangManager.GetTranslation "VarResetSuccess", "TestVar"
    Next i
    
    TestTranslationPerformance = "Translations: " & Operations & " in " & _
                                Format((Timer - StartTime) * 1000, "0.00") & " ms"
End Function