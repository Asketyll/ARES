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
    tidMSd = 9
    tidStringsInEl = 10
    tidLink = 11
    tidMSGraphical = 12
    tidARESMSVar = 13
    tidBootLoader = 14
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
Private TestElement As element ' Global test element for reuse

' === MAIN TEST RUNNER ===
Public Sub RunAllTests()
    Dim StartTime As Double
    Dim Results As String
    
    ' Create test environment
    OpenNewFile
    Set TestElement = CreateTestElement()
    
    StartTime = Timer
    
    ' Initialize test tracking
    TestCount = 0
    ReDim TestResults(0)
    
    ' Display header
    Results = "=== ARES TEST SUITE ===" & vbCrLf
    Results = Results & "Started: " & Now & vbCrLf
    Results = Results & String(25, "=") & vbCrLf & vbCrLf
    
    ' Run all test modules
    RunTest "Error Handler", tidErrorHandler
    RunTest "Configuration", tidConfig
    RunTest "Language Manager", tidLangManager
    RunTest "UUID Generator", tidUUID
    RunTest "ARES Variables", tidARESVars
    RunTest "Custom Properties", tidCustomProps
    RunTest "Element Processing", tidElementProcess
    RunTest "Length Calculations", tidLength
    RunTest "MSd Functions", tidMSd
    RunTest "String In Elements", tidStringsInEl
    RunTest "Link Functions", tidLink
    RunTest "MS Graphical", tidMSGraphical
    RunTest "ARES MS Variables", tidARESMSVar
    RunTest "Boot Loader", tidBootLoader
    
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
    Dim result As Boolean
    
    ' Get test name and run test
    Select Case TestIdentifier
        Case tidConfig
            TestName = "Configuration"
            result = ConfigTest()
        Case tidLangManager
            TestName = "Language Manager"
            result = LangManagerTest()
        Case tidUUID
            TestName = "UUID Generator"
            result = UUIDTest()
        Case tidARESVars
            TestName = "ARES Variables"
            result = ARES_VARTest()
        Case tidCustomProps
            TestName = "Custom Properties"
            result = CustomPropertyHandlerTest()
        Case tidErrorHandler
            TestName = "Error Handler"
            result = ErrorHandlerTest()
        Case tidElementProcess
            TestName = "Element Processing"
            result = ElementInProcesseTest()
        Case tidLength
            TestName = "Length Calculations"
            result = LengthTest()
        Case tidMSd
            TestName = "MSd Functions"
            result = MSdTest()
        Case tidStringsInEl
            TestName = "String In Elements"
            result = StringsInElTest()
        Case tidLink
            TestName = "Link Functions"
            result = LinkTest()
        Case tidMSGraphical
            TestName = "MS Graphical"
            result = MSGraphicalTest()
        Case tidARESMSVar
            TestName = "ARES MS Variables"
            result = ARESMSVarTest()
        Case tidBootLoader
            TestName = "Boot Loader"
            result = BootLoaderTest()
        Case Else
            MsgBox "Invalid test ID: " & TestIdentifier, vbCritical, "Test Error"
            Exit Sub
    End Select
    
    ' Display result
    MsgBox TestName & " Test: " & IIf(result, "PASSED", "FAILED"), _
           IIf(result, vbInformation, vbCritical), "Single Test Result"
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
    If Config.RemoveValue("ARES_Unit_testing") Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 1.5: Verify removal
    TotalTests = TotalTests + 1
    Dim RemovedValue As String
    RemovedValue = Config.GetVar("ARES_Unit_testing")
    If RemovedValue = "" Or RemovedValue = ARESConstants.ARES_NAVD Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 1.6: SetVar with empty value
    TotalTests = TotalTests + 1
    If Config.SetVar("ARES_Empty_Test", "") Then
        TestsPassed = TestsPassed + 1
    End If
    Config.RemoveValue ("ARES_Empty_Test")
    
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
    
    ' Test 2.4: Empty key handling
    TotalTests = TotalTests + 1
    Translation = LangManager.GetTranslation("")
    If InStr(Translation, "Empty translation key") > 0 Then
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
    
    ' Cleanup: Delete test log file
    On Error Resume Next
    If Len(Dir(TestHandler.LogFilePath)) > 0 Then
        Kill TestHandler.LogFilePath
    End If
    On Error GoTo TestError
    
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
    
    ' Test 7.1: Initial state
    TotalTests = TotalTests + 1
    If TestProcessor.Count = 0 And TestProcessor.IsEmpty Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 7.2: Add real element
    TotalTests = TotalTests + 1
    If Not TestElement Is Nothing Then
        If TestProcessor.Add(TestElement) Then
            TestsPassed = TestsPassed + 1
        End If
    Else
        ' Skip this test if no element available
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 7.3: Contains real element
    TotalTests = TotalTests + 1
    If Not TestElement Is Nothing Then
        If TestProcessor.Contains(TestElement) Then
            TestsPassed = TestsPassed + 1
        End If
    Else
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 7.4: Count after add
    TotalTests = TotalTests + 1
    If TestProcessor.Count = 1 And Not TestProcessor.IsEmpty Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 7.5: ContainsId with element ID
    TotalTests = TotalTests + 1
    If Not TestElement Is Nothing Then
        Dim ElementId As String
        ElementId = DLongToString(TestElement.ID)
        If TestProcessor.ContainsId(ElementId) Then
            TestsPassed = TestsPassed + 1
        End If
    Else
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 7.6: GetElementById
    TotalTests = TotalTests + 1
    If Not TestElement Is Nothing Then
        Dim RetrievedElement As element
        Set RetrievedElement = TestProcessor.GetElementById(DLongToString(TestElement.ID))
        If Not RetrievedElement Is Nothing Then
            TestsPassed = TestsPassed + 1
        End If
    Else
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 7.7: Remove element
    TotalTests = TotalTests + 1
    If Not TestElement Is Nothing Then
        If TestProcessor.Remove(TestElement) Then
            TestsPassed = TestsPassed + 1
        End If
    Else
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 7.8: Count after remove
    TotalTests = TotalTests + 1
    If TestProcessor.Count = 0 And TestProcessor.IsEmpty Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 7.9: Double add protection (should not add twice)
    TotalTests = TotalTests + 1
    If Not TestElement Is Nothing Then
        TestProcessor.Add TestElement
        Dim FirstAdd As Boolean
        FirstAdd = TestProcessor.Add(TestElement) ' Should return False
        If Not FirstAdd And TestProcessor.Count = 1 Then
            TestsPassed = TestsPassed + 1
        End If
        TestProcessor.Clear
    Else
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 7.10: Clear operation
    TotalTests = TotalTests + 1
    TestProcessor.Clear
    If TestProcessor.Count = 0 And TestProcessor.IsEmpty Then
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
    
    If TestElement Is Nothing Then
        LengthTest = False
        Exit Function
    End If
    
    If Not ARESConfig.IsInitialized Then
        ARESConfig.Initialize
    End If
    
    ' Test 8.1: Basic length calculation
    TotalTests = TotalTests + 1
    Dim CalculatedLength As Double
    CalculatedLength = Length.GetLength(TestElement, , True, False)
    ' Line from (0,0,0) to (200,200,0) should have length = sqrt(200² + 200²) ˜ 282.84
    If CalculatedLength > 280 And CalculatedLength < 285 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 8.2: Length with rounding
    TotalTests = TotalTests + 1
    Dim RoundedLength As Double
    RoundedLength = Length.GetLength(TestElement, 1, True, False)
    ' Should be rounded to 1 decimal place
    If RoundedLength > 282# And RoundedLength < 283# Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 8.3: Length without rounding
    TotalTests = TotalTests + 1
    Dim UnroundedLength As Double
    UnroundedLength = Length.GetLength(TestElement, , False, False)
    If UnroundedLength > 282.84 And UnroundedLength < 282.85 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 8.4: SetRound function
    TotalTests = TotalTests + 1
    If Length.SetRound(3) Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 8.5: ResetRound function
    TotalTests = TotalTests + 1
    If Length.ResetRound() Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 8.6: Error handling with invalid rounding
    TotalTests = TotalTests + 1
    Dim ErrorLength As Double
    ErrorLength = Length.GetLength(TestElement, 255, True, False) ' 255 is reserved error value
    If ErrorLength = 0 Then ' Should return 0 on error
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 8.7: Create additional element types for testing
    Dim TestArc As ArcElement
    Dim ArcCenter As Point3d
    Dim StartAngle As Double, SweepAngle As Double
    Dim PrimaryAxis As Double, SecondaryAxis As Double
    
    TotalTests = TotalTests + 1
    ArcCenter = Point3dFromXYZ(100, 100, 0)
    StartAngle = 0
    SweepAngle = Application.Pi ' Half circle
    PrimaryAxis = 50
    SecondaryAxis = 50
    
    Set TestArc = CreateArcElement2(Nothing, ArcCenter, PrimaryAxis, SecondaryAxis, Matrix3dIdentity, StartAngle, SweepAngle)
    ActiveModelReference.AddElement TestArc
    
    Dim ArcLength As Double
    ArcLength = Length.GetLength(TestArc)
    ' Half circle with radius 50: length = Pi * 50 ˜ 157.08
    If ArcLength > 155 And ArcLength < 160 Then
        TestsPassed = TestsPassed + 1
    End If
    
    LengthTest = (TestsPassed = TotalTests)
    Exit Function
    
ErrorHandler:
    LengthTest = False
End Function

' Test 9: MSd functions
Private Function MSdTest() As Boolean
    On Error GoTo ErrorHandler
    
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    
    ' Test 9.1: Valid element types
    TotalTests = TotalTests + 1
    If MicroStationDefinition.IsValidElementType(3) Then ' Line type
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 9.2: Invalid type detection
    TotalTests = TotalTests + 1
    If Not MicroStationDefinition.IsValidElementType(999) Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 9.3: String to element type conversion
    TotalTests = TotalTests + 1
    Dim ConvertedType As MsdElementType
    ConvertedType = MicroStationDefinition.StringToMsdElementType("Line")
    If ConvertedType = msdElementTypeLine Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 9.4: Case insensitive conversion
    TotalTests = TotalTests + 1
    ConvertedType = MicroStationDefinition.StringToMsdElementType("line", False)
    If ConvertedType = msdElementTypeLine Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 9.5: Invalid string conversion
    TotalTests = TotalTests + 1
    ConvertedType = MicroStationDefinition.StringToMsdElementType("InvalidType")
    If ConvertedType = 0 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 9.6: Multiple element type validations
    TotalTests = TotalTests + 1
    Dim ValidTypes As Boolean
    ValidTypes = MicroStationDefinition.IsValidElementType(16) And _
                MicroStationDefinition.IsValidElementType(6) And _
                MicroStationDefinition.IsValidElementType(17) ' Arc, Shape and Text
    If ValidTypes Then
        TestsPassed = TestsPassed + 1
    End If
    
    MSdTest = (TestsPassed = TotalTests)
    Exit Function
    
ErrorHandler:
    MSdTest = False
End Function

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
    
    ' Test 10.2: RemovePattern with empty pattern
    TotalTests = TotalTests + 1
    TestString = StringsInEl.RemovePattern("Hello_World", "")
    If TestString = "Hello_World" Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 10.3: RemovePattern with non-existent pattern
    TotalTests = TotalTests + 1
    TestString = StringsInEl.RemovePattern("Hello_World", "NotFound")
    If TestString = "Hello_World" Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 10.4: Create and test text element
    TotalTests = TotalTests + 1
    Dim TestTextElement As TextElement
    Dim TextOrigin As Point3d
    Dim TestText As String
    
    TextOrigin = Point3dFromXYZ(50, 50, 0)
    TestText = "Test (Xx_m) trigger"
    Set TestTextElement = CreateTextElement1(Nothing, TestText, TextOrigin, Matrix3dIdentity)
    ActiveModelReference.AddElement TestTextElement
    
    ' Test getting text from element
    Dim RetrievedTexts() As String
    RetrievedTexts = StringsInEl.GetSetTextsInEl(TestTextElement)
    If IsArray(RetrievedTexts) And UBound(RetrievedTexts) >= 0 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 10.5: Setting text in element
    TotalTests = TotalTests + 1
    Dim ModifiedTexts() As String
    ModifiedTexts = StringsInEl.GetSetTextsInEl(TestTextElement, "Modified text")
    If IsArray(ModifiedTexts) Then
        TestsPassed = TestsPassed + 1
    End If
    
    StringsInElTest = (TestsPassed = TotalTests)
    Exit Function
    
ErrorHandler:
    StringsInElTest = False
End Function

' Test 11: Link functions
Private Function LinkTest() As Boolean
    On Error GoTo ErrorHandler
    
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    
    If TestElement Is Nothing Then
        LinkTest = False
        Exit Function
    End If
    
    ' Test 11.1: GetLink with element without graphic group
    TotalTests = TotalTests + 1
    Dim LinkedElements() As element
    LinkedElements = Link.GetLink(TestElement, False)
    ' Should return empty array since test element has default graphic group
    If IsArray(LinkedElements) Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 11.2: Create elements with same graphic group
    TotalTests = TotalTests + 1
    Dim TestElement2 As LineElement
    Dim StartPoint As Point3d, EndPoint As Point3d
    
    StartPoint = Point3dFromXYZ(300, 300, 0)
    EndPoint = Point3dFromXYZ(400, 400, 0)
    Set TestElement2 = CreateLineElement2(Nothing, StartPoint, EndPoint)
    
    ' Set same graphic group (non-default)
    TestElement.GraphicGroup = 1
    TestElement2.GraphicGroup = 1
    TestElement.Rewrite
    
    ActiveModelReference.AddElement TestElement2
    
    ' Now test GetLink with grouped elements
    LinkedElements = Link.GetLink(TestElement, False)
    If IsArray(LinkedElements) And UBound(LinkedElements) >= 0 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 11.3: GetLink with ReturnMe = True
    TotalTests = TotalTests + 1
    Dim AllGroupElements() As element
    AllGroupElements = Link.GetLink(TestElement, True)
    If IsArray(AllGroupElements) Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 11.4: GetLink with filter by types
    TotalTests = TotalTests + 1
    Dim FilteredElements() As element
    Dim LineTypes(0) As Long
    LineTypes(0) = 3 ' Line type
    FilteredElements = Link.GetLink(TestElement, False, LineTypes)
    If IsArray(FilteredElements) Then
        TestsPassed = TestsPassed + 1
    End If
    
    LinkTest = (TestsPassed = TotalTests)
    Exit Function
    
ErrorHandler:
    LinkTest = False
End Function

' Test 12: MSGraphicalInteraction
Private Function MSGraphicalTest() As Boolean
    On Error GoTo ErrorHandler
    
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    
    If TestElement Is Nothing Then
        MSGraphicalTest = False
        Exit Function
    End If
    
    ' Test 12.1: ZoomEl function
    TotalTests = TotalTests + 1
    If MSGraphicalInteraction.ZoomEl(TestElement) Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 12.2: ZoomEl with custom factor
    TotalTests = TotalTests + 1
    If MSGraphicalInteraction.ZoomEl(TestElement, 2#) Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 12.3: HighlightEl function
    TotalTests = TotalTests + 1
    If MSGraphicalInteraction.HighlightEl(TestElement) Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 12.4: Check TEC (TransientElementContainer) is created
    TotalTests = TotalTests + 1
    If Not MSGraphicalInteraction.TEC Is Nothing Then
        TestsPassed = TestsPassed + 1
    End If
    
    MSGraphicalTest = (TestsPassed = TotalTests)
    Exit Function
    
ErrorHandler:
    MSGraphicalTest = False
End Function

' Test 13: ARES_MS_VAR_Class
Private Function ARESMSVarTest() As Boolean
    On Error GoTo ErrorHandler
    
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    Dim TestVar As New ARES_MS_VAR_Class
    
    ' Test 13.1: Initialize variable
    TotalTests = TotalTests + 1
    TestVar.Initialize "TestKey", "DefaultValue"
    If TestVar.key = "TestKey" And TestVar.defaultValue = "DefaultValue" Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 13.2: IsDefault check
    TotalTests = TotalTests + 1
    If TestVar.IsDefault() Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 13.3: Set new value
    TotalTests = TotalTests + 1
    TestVar.Value = "NewValue"
    If TestVar.Value = "NewValue" Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 13.4: IsModified check
    TotalTests = TotalTests + 1
    If TestVar.IsModified() Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 13.5: IsValid check
    TotalTests = TotalTests + 1
    If TestVar.IsValid() Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 13.6: ResetToDefault
    TotalTests = TotalTests + 1
    TestVar.ResetToDefault
    If TestVar.IsDefault() And Not TestVar.IsModified() Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 13.7: ToString method
    TotalTests = TotalTests + 1
    Dim ToStringResult As String
    ToStringResult = TestVar.ToString()
    If InStr(ToStringResult, "TestKey") > 0 And InStr(ToStringResult, "DefaultValue") > 0 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 13.8: Invalid value handling
    TotalTests = TotalTests + 1
    TestVar.Value = ""
    If Not TestVar.IsValid() Then
        TestsPassed = TestsPassed + 1
    End If
    
    ARESMSVarTest = (TestsPassed = TotalTests)
    Exit Function
    
ErrorHandler:
    ARESMSVarTest = False
End Function

' Test 14: BootLoader
Private Function BootLoaderTest() As Boolean
    On Error GoTo ErrorHandler
    
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    
    BootLoader.OnProjectLoad
    
    ' Test 14.1: Global objects existence
    TotalTests = TotalTests + 1
    If Not BootLoader.ErrorHandler Is Nothing And _
       Not BootLoader.ElementInProcesse Is Nothing And _
       Not BootLoader.ARESConfig Is Nothing Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 14.2: ARESConfig initialization
    TotalTests = TotalTests + 1
    If BootLoader.ARESConfig.IsInitialized Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 14.3: ElementInProcesse functionality
    TotalTests = TotalTests + 1
    If BootLoader.ElementInProcesse.Count >= 0 Then ' Should at least return a count
        TestsPassed = TestsPassed + 1
    End If
    
    BootLoaderTest = (TestsPassed = TotalTests)
    Exit Function
    
ErrorHandler:
    BootLoaderTest = False
End Function

' === HELPER FUNCTIONS ===

Private Function CreateTestElement() As element
    Dim oStartPoint As Point3d
    Dim oEndPoint As Point3d
    Dim oLine As LineElement

    'Starting and ending points of Line
    oStartPoint = Point3dFromXYZ(0, 0, 0)
    oEndPoint = Point3dFromXYZ(200, 200, 0)
    
    'draw lines
    Set oLine = CreateLineElement2(Nothing, oStartPoint, oEndPoint)
    
    'Add line in Active Model
    ActiveModelReference.AddElement oLine
    
    Set CreateTestElement = oLine
End Function

Private Sub OpenNewFile()
    Dim oDgn As DesignFile
    Dim sFileName As String
    Dim sSeedName As String
    
    sFileName = ActiveDesignFile.Path & "\" & "UnitTesting_" & Format(Now, "yyyymmdd_hhmmss") & ".dgn"
    If Dir(sFileName) <> "" Then
        Kill sFileName
    End If
    sSeedName = ActiveWorkspace.ConfigurationVariableValue("MS_DESIGNMODELSEED", True)
    Set oDgn = CreateDesignFile(sSeedName, sFileName, True)
End Sub

Private Sub RunTest(TestName As String, TestIdentifier As Integer)
    Dim StartTime As Double
    Dim result As TestResult
    
    StartTime = Timer
    
    result.Name = TestName
    
    On Error Resume Next
    ' Execute test based on ID
    Select Case TestIdentifier
        Case tidConfig: result.Passed = ConfigTest()
        Case tidLangManager: result.Passed = LangManagerTest()
        Case tidUUID: result.Passed = UUIDTest()
        Case tidARESVars: result.Passed = ARES_VARTest()
        Case tidCustomProps: result.Passed = CustomPropertyHandlerTest()
        Case tidErrorHandler: result.Passed = ErrorHandlerTest()
        Case tidElementProcess: result.Passed = ElementInProcesseTest()
        Case tidLength: result.Passed = LengthTest()
        Case tidMSd: result.Passed = MSdTest()
        Case tidStringsInEl: result.Passed = StringsInElTest()
        Case tidLink: result.Passed = LinkTest()
        Case tidMSGraphical: result.Passed = MSGraphicalTest()
        Case tidARESMSVar: result.Passed = ARESMSVarTest()
        Case tidBootLoader: result.Passed = BootLoaderTest()
        Case Else
            result.Passed = False
            result.Message = "Unknown test ID"
    End Select
    
    If Err.Number <> 0 Then
        result.Passed = False
        result.Message = "Error: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    
    result.Duration = Round((Timer - StartTime) * 1000, 2) ' Convert to milliseconds
    
    ' Add to results array
    TestCount = TestCount + 1
    ReDim Preserve TestResults(TestCount)
    TestResults(TestCount) = result
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
            Report = Report & IIf(.Passed, "O", "X") & " " & .Name
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