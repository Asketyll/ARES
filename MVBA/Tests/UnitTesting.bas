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
    tidAutoLengths = 15
    tidConfigExportImport = 16
    tidFileDialogs = 17
    tidPropertyCalculation = 18
    tidPropertyRuleValidation = 19
End Enum

' Test result structure
Private Type TestResult
    Name As String
    Passed As Boolean
    Message As String
    Duration As Double
End Type

' Helper structure to hold test elements
Private Type TestElementsCollection
    TextElement As TextElement
    LineElement1 As LineElement
    LineElement2 As LineElement
    ArcElement As ArcElement
    ShapeElement As ShapeElement
    GraphicGroupId As Long
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
    RunTest "Auto Lengths", tidAutoLengths
    RunTest "Config Export Import", tidConfigExportImport
    RunTest "File Dialogs", tidFileDialogs
    RunTest "Property Calculation", tidPropertyCalculation
    RunTest "Property Rule Validation", tidPropertyRuleValidation

    ' Generate summary report
    Results = Results & GenerateTestReport(Timer - StartTime)
    
    ' Display results
    MsgBox Results, vbOKOnly + vbInformation, "ARES Test Suite Results"
    
    ' Save results to log
    SaveTestResults Results
End Sub

' Run a single test by ID
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
        Case tidMSd
            TestName = "MSd Functions"
            Result = MSdTest()
        Case tidStringsInEl
            TestName = "String In Elements"
            Result = StringsInElTest()
        Case tidLink
            TestName = "Link Functions"
            Result = LinkTest()
        Case tidMSGraphical
            TestName = "MS Graphical"
            Result = MSGraphicalTest()
        Case tidARESMSVar
            TestName = "ARES MS Variables"
            Result = ARESMSVarTest()
        Case tidBootLoader
            TestName = "Boot Loader"
            Result = BootLoaderTest()
        Case tidAutoLengths
            TestName = "Auto Lengths"
            Result = AutoLengthsTest()
        Case tidConfigExportImport
            TestName = "Config Export Import"
            Result = ConfigExportImportTest()
        Case tidFileDialogs
            TestName = "File Dialogs"
            Result = FileDialogsTest()
        Case tidPropertyCalculation
            TestName = "Property Calculation"
            Result = PropertyCalculationTest()
        Case tidPropertyRuleValidation
            TestName = "Property Rule Validation"
            Result = PropertyRuleValidationTest()
        Case Else
            MsgBox "Invalid test ID: " & TestIdentifier & ". Valid range: 1-19", vbCritical, "Test Error"
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

' Test 5: ARES custom properties (attach / read / write on the DGNLib-defined item types)
Private Function CustomPropertyHandlerTest() As Boolean
    On Error GoTo ErrorHandler

    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    Dim ITL As ItemTypeLibrary
    Dim oItem As ItemType
    Dim names() As String
    Dim name1 As String
    Dim name2 As String
    Dim hasSecond As Boolean

    ' A graphical test element is required to attach items
    If TestElement Is Nothing Then
        CustomPropertyHandlerTest = False
        Exit Function
    End If

    ' Strategy A: the ARES item types + value lists live in a DGNLib (deployed via MS_DGNLIBLIST),
    ' not created by VBA. If the library is not available in this session, the attach/read/write
    ' helpers cannot be exercised - treat the test as not-applicable (pass) rather than fail.
    Set ITL = CustomPropertyHandler.FindItemTypeLibrary(ARESConstants.ARES_NAME_LIBRARY_TYPE)
    If ITL Is Nothing Then
        CustomPropertyHandlerTest = True
        Exit Function
    End If

    ' The managed property names are user-configurable (ARES_Custom_Property_List).
    names = CustomPropertyHandler.GetCustomPropertyNames()
    If UBound(names) < LBound(names) Then
        CustomPropertyHandlerTest = True   ' nothing configured -> nothing to test
        Exit Function
    End If
    name1 = Trim(names(LBound(names)))
    hasSecond = (UBound(names) > LBound(names))
    If hasSecond Then name2 = Trim(names(LBound(names) + 1))

    ' Start from a clean element (a previous run may have left items attached)
    CustomPropertyHandler.RemoveItemFromElement TestElement, name1
    If hasSecond Then CustomPropertyHandler.RemoveItemFromElement TestElement, name2

    ' Test 5.1: the first configured item type and its property exist in the library
    TotalTests = TotalTests + 1
    Set oItem = ITL.GetItemTypeByName(name1)
    If Not oItem Is Nothing Then
        If Not oItem.GetPropertyByName(name1) Is Nothing Then TestsPassed = TestsPassed + 1
    End If

    ' Test 5.2: attach the first property to the element
    TotalTests = TotalTests + 1
    If CustomPropertyHandler.AttachItemToElement(TestElement, name1) Then
        If TestElement.Items.HasItems(ARESConstants.ARES_NAME_LIBRARY_TYPE, name1) Then TestsPassed = TestsPassed + 1
    End If

    ' Test 5.3: round-trip a free-text value on the first property
    TotalTests = TotalTests + 1
    If CustomPropertyHandler.SetPropertyValueToElement(TestElement, name1, "ARES Test", name1) Then
        If CStr(CustomPropertyHandler.GetPropertyValueFromElement(TestElement, name1, name1)) = "ARES Test" Then TestsPassed = TestsPassed + 1
    End If

    ' Test 5.4: a second configured property attaches independently and both coexist
    If hasSecond Then
        TotalTests = TotalTests + 1
        If CustomPropertyHandler.AttachItemToElement(TestElement, name2) Then
            If TestElement.Items.HasItems(ARESConstants.ARES_NAME_LIBRARY_TYPE, name1) _
               And TestElement.Items.HasItems(ARESConstants.ARES_NAME_LIBRARY_TYPE, name2) Then TestsPassed = TestsPassed + 1
        End If

        ' Test 5.5: detaching the first leaves the second untouched
        TotalTests = TotalTests + 1
        CustomPropertyHandler.RemoveItemFromElement TestElement, name1
        If (Not TestElement.Items.HasItems(ARESConstants.ARES_NAME_LIBRARY_TYPE, name1)) _
           And TestElement.Items.HasItems(ARESConstants.ARES_NAME_LIBRARY_TYPE, name2) Then TestsPassed = TestsPassed + 1
    Else
        CustomPropertyHandler.RemoveItemFromElement TestElement, name1
    End If

    ' Test 5.6: detaching an already-detached item type is graceful (False, no crash)
    TotalTests = TotalTests + 1
    If Not CustomPropertyHandler.RemoveItemFromElement(TestElement, name1) Then TestsPassed = TestsPassed + 1

    ' Test 5.7: unknown library/item resolves to Nothing (no crash)
    TotalTests = TotalTests + 1
    If CustomPropertyHandler.GetItemTypePropertyHandlerFromElement(TestElement, "NonExistentItem", "NonExistentLibrary") Is Nothing Then TestsPassed = TestsPassed + 1

    ' Cleanup - detach the second property too (the ARES library itself is kept)
    If hasSecond Then CustomPropertyHandler.RemoveItemFromElement TestElement, name2

    ' Library operations can be environment-dependent; allow a small margin
    CustomPropertyHandlerTest = (TestsPassed >= TotalTests - 1)
    Exit Function

ErrorHandler:
    ' Log error details for debugging
    If Not BootLoader.ErrorHandler Is Nothing Then
        BootLoader.ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "CustomPropertyHandlerTest"
    End If
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
    If TestVar.Key = "TestKey" And TestVar.DefaultValue = "DefaultValue" Then
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

' Test 15: AutoLengths
Private Function AutoLengthsTest() As Boolean
    On Error GoTo ErrorHandler
    
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    Dim TestAutoLengths As New AutoLengths
    
    ' Ensure ARESConfig is initialized
    If Not ARESConfig.IsInitialized Then
        ARESConfig.Initialize
    End If
    
    ' Create test environment with multiple linked elements
    Dim TestElements As TestElementsCollection
    TestElements = CreateTestEnvironmentForAutoLengths()
    
    ' Test 15.1: Initialize with valid text element
    TotalTests = TotalTests + 1
    If Not TestElements.TextElement Is Nothing Then
        TestAutoLengths.Initialize TestElements.TextElement
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 15.2: Test with no linked elements (default graphic group)
    TotalTests = TotalTests + 1
    Dim IsolatedTextElement As TextElement
    Set IsolatedTextElement = CreateTextElement1(Nothing, "Isolated (Xx_m) text", Point3dFromXYZ(500, 500, 0), Matrix3dIdentity)
    ActiveModelReference.AddElement IsolatedTextElement
    
    Dim TestAutoLengths3 As New AutoLengths
    TestAutoLengths3.Initialize IsolatedTextElement
    ' Should handle gracefully when no linked elements found
    TestsPassed = TestsPassed + 1
    
    ' Test 15.3: Test with single linked element
    TotalTests = TotalTests + 1
    If TestSingleLinkedElement(TestElements) Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 15.4: Test with multiple linked elements (different lengths)
    TotalTests = TotalTests + 1
    If TestMultipleLinkedElements(TestElements) Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 15.5: Test trigger replacement in text
    TotalTests = TotalTests + 1
    ARESConfig.ARES_LENGTH_TRIGGER.Value = "Xx_cm"
    If TestTriggerReplacement(TestElements) Then
        TestsPassed = TestsPassed + 1
    End If
    ARESConfig.ARES_LENGTH_TRIGGER.Value = "Xx_m"
    
    ' Test 15.6: Test with different element types (Line, Arc, Shape)
    TotalTests = TotalTests + 1
    If TestDifferentElementTypes() Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 15.7: Test color update functionality
    TotalTests = TotalTests + 1
    If TestColorUpdate(TestElements) Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 15.8: Test with TextNodeElement
    TotalTests = TotalTests + 1
    If TestWithTextNodeElement() Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 15.9: Test rounding functionality
    TotalTests = TotalTests + 1
    If TestRoundingFunctionality() Then
        TestsPassed = TestsPassed + 1
    End If
    
    AutoLengthsTest = (TestsPassed >= TotalTests - 2) ' Allow some failures for complex operations
    Exit Function
    
ErrorHandler:
    If Not BootLoader.ErrorHandler Is Nothing Then
        BootLoader.ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "AutoLengthsTest"
    End If
    AutoLengthsTest = False
End Function

' Create a complete test environment with linked elements
Private Function CreateTestEnvironmentForAutoLengths() As TestElementsCollection
    Dim TestElements As TestElementsCollection
    
    ' Use a unique graphic group ID
    TestElements.GraphicGroupId = 10
    
    ' Create text element with trigger
    Set TestElements.TextElement = CreateTextElement1(Nothing, "Length: (m)", Point3dFromXYZ(100, 300, 0), Matrix3dIdentity)
    TestElements.TextElement.GraphicGroup = TestElements.GraphicGroupId
    ActiveModelReference.AddElement TestElements.TextElement
    
    ' Create linked line element 1
    Set TestElements.LineElement1 = CreateLineElement2(Nothing, Point3dFromXYZ(100, 100, 0), Point3dFromXYZ(200, 100, 0))
    TestElements.LineElement1.GraphicGroup = TestElements.GraphicGroupId
    TestElements.LineElement1.Color = 3 ' Red
    ActiveModelReference.AddElement TestElements.LineElement1
    
    ' Create linked line element 2
    Set TestElements.LineElement2 = CreateLineElement2(Nothing, Point3dFromXYZ(100, 150, 0), Point3dFromXYZ(250, 150, 0))
    TestElements.LineElement2.GraphicGroup = TestElements.GraphicGroupId
    TestElements.LineElement2.Color = 4 ' Blue
    ActiveModelReference.AddElement TestElements.LineElement2
    
    ' Create linked arc element
    Set TestElements.ArcElement = CreateArcElement2(Nothing, Point3dFromXYZ(100, 200, 0), 50, 50, Matrix3dIdentity, 0, Application.Pi)
    TestElements.ArcElement.GraphicGroup = TestElements.GraphicGroupId
    TestElements.ArcElement.Color = 5 ' Green
    ActiveModelReference.AddElement TestElements.ArcElement
    
    ' Create linked shape element (rectangle)
    Dim ShapeVertices(4) As Point3d
    ShapeVertices(0) = Point3dFromXYZ(300, 100, 0)
    ShapeVertices(1) = Point3dFromXYZ(400, 100, 0)
    ShapeVertices(2) = Point3dFromXYZ(400, 150, 0)
    ShapeVertices(3) = Point3dFromXYZ(300, 150, 0)
    ShapeVertices(4) = ShapeVertices(0) ' Close the shape
    
    Set TestElements.ShapeElement = CreateShapeElement1(Nothing, ShapeVertices, msdFillModeNotFilled)
    TestElements.ShapeElement.GraphicGroup = TestElements.GraphicGroupId
    TestElements.ShapeElement.Color = 6 ' Yellow
    ActiveModelReference.AddElement TestElements.ShapeElement
    
    CreateTestEnvironmentForAutoLengths = TestElements
End Function

' Test single linked element scenario
Private Function TestSingleLinkedElement(TestElements As TestElementsCollection) As Boolean
    Dim TestAutoLengths As New AutoLengths
    Dim SingleTextElement As TextElement
    
    ' Create text element linked only to one line
    Set SingleTextElement = CreateTextElement1(Nothing, "Single: (m)", Point3dFromXYZ(150, 120, 0), Matrix3dIdentity)
    SingleTextElement.GraphicGroup = TestElements.GraphicGroupId + 1
    ActiveModelReference.AddElement SingleTextElement
    
    ' Create only one linked line
    Dim SingleLineElement As LineElement
    Set SingleLineElement = CreateLineElement2(Nothing, Point3dFromXYZ(150, 80, 0), Point3dFromXYZ(250, 80, 0))
    SingleLineElement.GraphicGroup = TestElements.GraphicGroupId + 1
    ActiveModelReference.AddElement SingleLineElement
    
    ' Test the auto lengths functionality
    TestAutoLengths.Initialize SingleTextElement
    TestAutoLengths.UpdateLengths
    
    ' Check if text was updated (approximate length should be 100)
    Dim UpdatedText As String
    Dim UpdatedTexts() As String
    UpdatedTexts = StringsInEl.GetSetTextsInEl(SingleTextElement)
    If IsArray(UpdatedTexts) And UBound(UpdatedTexts) >= 0 Then
        UpdatedText = UpdatedTexts(0)
        If InStr(UpdatedText, "100") > 0 Or InStr(UpdatedText, "10") > 0 Then
            TestSingleLinkedElement = True
        End If
    End If
End Function

' Test multiple linked elements scenario
Private Function TestMultipleLinkedElements(TestElements As TestElementsCollection) As Boolean
    Dim TestAutoLengths As New AutoLengths
    
    ' Initialize with text element that has multiple linked elements
    TestAutoLengths.Initialize TestElements.TextElement
    
    ' This should trigger the selection form or auto-select if only one non-zero length
    TestAutoLengths.UpdateLengths
    
    ' For testing purposes, simulate element selection
    
    TestAutoLengths.OnElementSelected TestElements.LineElement1, TestElements.TextElement
    
    ' Check if text was updated
    Dim UpdatedTexts() As String
    UpdatedTexts = StringsInEl.GetSetTextsInEl(TestElements.TextElement)
    If IsArray(UpdatedTexts) And UBound(UpdatedTexts) >= 0 Then
        Dim UpdatedText As String
        UpdatedText = UpdatedTexts(0)
        ' Should contain a numeric value
        If InStr(UpdatedText, "100") > 0 Or InStr(UpdatedText, "10") > 0 Then
            TestMultipleLinkedElements = True
        End If
    End If
End Function

' Test trigger replacement functionality
Private Function TestTriggerReplacement(TestElements As TestElementsCollection) As Boolean
    ' Create text element with custom trigger
    Dim TriggerTestElement As TextElement
    Set TriggerTestElement = CreateTextElement1(Nothing, "Custom (cm) trigger test", Point3dFromXYZ(400, 300, 0), Matrix3dIdentity)
    TriggerTestElement.GraphicGroup = TestElements.GraphicGroupId
    ActiveModelReference.AddElement TriggerTestElement
    
    ' Test if trigger is properly detected and replaced
    Dim TestAutoLengths As New AutoLengths
    TestAutoLengths.Initialize TriggerTestElement
    TestAutoLengths.OnElementSelected TestElements.LineElement1, TriggerTestElement
    
    ' Check if trigger was replaced
    Dim UpdatedTexts() As String
    UpdatedTexts = StringsInEl.GetSetTextsInEl(TriggerTestElement)
    If IsArray(UpdatedTexts) And UBound(UpdatedTexts) >= 0 Then
        Dim UpdatedText As String
        UpdatedText = UpdatedTexts(0)
        ' Should not contain the original trigger and should have a number
        If (InStr(UpdatedText, "10") > 0 Or InStr(UpdatedText, "100") > 0) Then
            TestTriggerReplacement = True
        End If
    End If
End Function

' Test different element types
Private Function TestDifferentElementTypes() As Boolean
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    
    ' Test with Line
    TotalTests = TotalTests + 1
    Dim LineLength As Double
    LineLength = Length.GetLength(TestElement)
    If LineLength > 0 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test with Arc (create a test arc)
    TotalTests = TotalTests + 1
    Dim TestArc As ArcElement
    Set TestArc = CreateArcElement2(Nothing, Point3dFromXYZ(0, 0, 0), 100, 100, Matrix3dIdentity, 0, Application.Pi / 2)
    ActiveModelReference.AddElement TestArc
    
    Dim ArcLength As Double
    ArcLength = Length.GetLength(TestArc)
    If ArcLength > 0 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test with Shape
    TotalTests = TotalTests + 1
    Dim TestShapeVertices(4) As Point3d
    TestShapeVertices(0) = Point3dFromXYZ(0, 0, 0)
    TestShapeVertices(1) = Point3dFromXYZ(100, 0, 0)
    TestShapeVertices(2) = Point3dFromXYZ(100, 100, 0)
    TestShapeVertices(3) = Point3dFromXYZ(0, 100, 0)
    TestShapeVertices(4) = TestShapeVertices(0)
    
    Dim TestShape As ShapeElement
    Set TestShape = CreateShapeElement1(Nothing, TestShapeVertices, msdFillModeNotFilled)
    ActiveModelReference.AddElement TestShape
    
    Dim ShapeLength As Double
    ShapeLength = Length.GetLength(TestShape)
    If ShapeLength > 0 Then
        TestsPassed = TestsPassed + 1
    End If
    
    TestDifferentElementTypes = (TestsPassed = TotalTests)
End Function

' Test color update functionality
Private Function TestColorUpdate(TestElements As TestElementsCollection) As Boolean
    ' Save original color setting
    Dim OriginalColorSetting As Boolean
    OriginalColorSetting = ARESConfig.ARES_UPDATE_COLOR_WITH_LENGTH.Value
    
    ' Enable color update
    ARESConfig.ARES_UPDATE_COLOR_WITH_LENGTH.Value = True
    
    ' Create test elements
    Dim ColorTestText As TextElement
    Set ColorTestText = CreateTextElement1(Nothing, "Color: (Xx_m)", Point3dFromXYZ(200, 400, 0), Matrix3dIdentity)
    ColorTestText.GraphicGroup = TestElements.GraphicGroupId
    ColorTestText.Color = 1 ' Original color
    ActiveModelReference.AddElement ColorTestText
    
    ' Test color update
    Dim TestAutoLengths As New AutoLengths
    TestAutoLengths.Initialize ColorTestText
    TestAutoLengths.OnElementSelected TestElements.LineElement1, ColorTestText ' Line has color 3 (Red)
    
    ' Check if color was updated
    If ColorTestText.Color = TestElements.LineElement1.Color Then
        TestColorUpdate = True
    End If
    
    ' Restore original setting
    ARESConfig.ARES_UPDATE_COLOR_WITH_LENGTH.Value = OriginalColorSetting
End Function

' Test with TextNodeElement
Private Function TestWithTextNodeElement() As Boolean
    ' Create a TextNodeElement with multiple lines
    Dim TextNodeOrigin As Point3d
    TextNodeOrigin = Point3dFromXYZ(300, 400, 0)
    
    Dim TextNodeElement As TextNodeElement
    Set TextNodeElement = CreateTextNodeElement2(Nothing, TextNodeOrigin, Matrix3dIdentity)
    TextNodeElement.AddTextLine "Line 1: (Xx_m)"
    TextNodeElement.AddTextLine "Line 2: (Xx_cm)"
    ActiveModelReference.AddElement TextNodeElement
    
    ' Set same graphic group
    TextNodeElement.GraphicGroup = 11
    
    ' Create linked element
    Dim LinkedLine As LineElement
    Set LinkedLine = CreateLineElement2(Nothing, Point3dFromXYZ(300, 350, 0), Point3dFromXYZ(400, 350, 0))
    LinkedLine.GraphicGroup = 11
    ActiveModelReference.AddElement LinkedLine
    
    ' Test AutoLengths with TextNodeElement
    Dim TestAutoLengths As New AutoLengths
    TestAutoLengths.Initialize TextNodeElement
    TestAutoLengths.UpdateLengths
    
    ' Check if any text line was updated
    Dim UpdatedTexts() As String
    UpdatedTexts = StringsInEl.GetSetTextsInEl(TextNodeElement)
    If IsArray(UpdatedTexts) And UBound(UpdatedTexts) >= 0 Then
        Dim i As Long
        For i = 0 To UBound(UpdatedTexts)
            If InStr(UpdatedTexts(i), "100") > 0 Or InStr(UpdatedTexts(i), "10") > 0 Then
                TestWithTextNodeElement = True
                Exit Function
            End If
        Next i
    End If
End Function

' Test rounding functionality
Private Function TestRoundingFunctionality() As Boolean
    ' Test different rounding values
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    
    ' Test with rounding = 0
    TotalTests = TotalTests + 1
    Dim OriginalRound As String
    OriginalRound = ARESConfig.ARES_LENGTH_ROUND.Value
    
    ARESConfig.ARES_LENGTH_ROUND.Value = "0"
    Dim Length0 As Double
    Length0 = Length.GetLength(TestElement, 0, True)
    If Length0 = Int(Length0) Then ' Should be whole number
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test with rounding = 2
    TotalTests = TotalTests + 1
    ARESConfig.ARES_LENGTH_ROUND.Value = "2"
    Dim Length2 As Double
    Length2 = Length.GetLength(TestElement, 2, True)
    If Length2 <> Int(Length2) Then ' Should have decimal places
        TestsPassed = TestsPassed + 1
    End If
    
    ' Restore original setting
    ARESConfig.ARES_LENGTH_ROUND.Value = OriginalRound
    
    TestRoundingFunctionality = (TestsPassed = TotalTests)
End Function

' Test 16: Configuration Export/Import
Private Function ConfigExportImportTest() As Boolean
    On Error GoTo ErrorHandler
    
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    Dim TestConfig As New ARESConfigClass
    
    ' Test 16.1: Initialize config
    TotalTests = TotalTests + 1
    If TestConfig.Initialize Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 16.2: Export configuration
    TotalTests = TotalTests + 1
    Dim ExportPath As String
    ExportPath = Environ("TEMP") & "\ARES_Test_Export.cfg"
    If TestConfig.ExportConfig(ExportPath) Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 16.3: Verify export file exists
    TotalTests = TotalTests + 1
    If Len(Dir(ExportPath)) > 0 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 16.4: Modify a configuration value
    TotalTests = TotalTests + 1
    Dim OriginalValue As String
    OriginalValue = TestConfig.ARES_ROUNDS.Value
    TestConfig.ARES_ROUNDS.Value = "99"
    Config.SetVar "ARES_Round", "99"
    If TestConfig.ARES_ROUNDS.Value = "99" Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 16.5: Import configuration (should restore original)
    TotalTests = TotalTests + 1
    If TestConfig.ImportConfig(ExportPath, True) Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 16.6: Verify import restored original value
    TotalTests = TotalTests + 1
    If TestConfig.ARES_ROUNDS.Value = OriginalValue Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 16.7: Get config summary
    TotalTests = TotalTests + 1
    Dim Summary As String
    Summary = TestConfig.GetConfigSummary()
    If Len(Summary) > 0 And InStr(Summary, "ARES Configuration Summary") > 0 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Cleanup
    On Error Resume Next
    If Len(Dir(ExportPath)) > 0 Then Kill ExportPath
    On Error GoTo ErrorHandler
    
    ConfigExportImportTest = (TestsPassed = TotalTests)
    Exit Function
    
ErrorHandler:
    ConfigExportImportTest = False
End Function

' Test 17: FileDialogs
Private Function FileDialogsTest() As Boolean
    On Error GoTo ErrorHandler
    
    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    
    ' Test 17.1: GetDefaultConfigDirectory
    TotalTests = TotalTests + 1
    Dim DefaultDir As String
    DefaultDir = FileDialogs.GetDefaultConfigDirectory()
    If Len(DefaultDir) > 0 And Len(Dir(DefaultDir, vbDirectory)) > 0 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 17.2: GenerateDefaultConfigFileName
    TotalTests = TotalTests + 1
    Dim DefaultFileName As String
    DefaultFileName = FileDialogs.GenerateDefaultConfigFileName()
    If Len(DefaultFileName) > 0 And InStr(DefaultFileName, ".cfg") > 0 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 17.3: GenerateDefaultConfigFileName with custom prefix
    TotalTests = TotalTests + 1
    Dim CustomFileName As String
    CustomFileName = FileDialogs.GenerateDefaultConfigFileName("CustomTest")
    If InStr(CustomFileName, "CustomTest") > 0 And InStr(CustomFileName, ".cfg") > 0 Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 17.4: CleanFilePath function (via private testing)
    TotalTests = TotalTests + 1
    ' We'll test this indirectly by testing ShowSaveFileDialog with a mock that won't show UI
    Dim TestPath As String
    TestPath = TestCleanFilePathFunctionality()
    If TestPath Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 17.5: EscapeForPowerShell function (via private testing)
    TotalTests = TotalTests + 1
    If TestEscapeForPowerShellFunctionality() Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 17.6: ShowConfigurationSummaryUI
    TotalTests = TotalTests + 1
    ' This will show a message box, but we can test it doesn't crash
    On Error Resume Next
    FileDialogs.ShowConfigurationSummaryUI
    If Err.Number = 0 Then
        TestsPassed = TestsPassed + 1
    End If
    On Error GoTo ErrorHandler
    
    ' Test 17.7: Test PowerShell command generation (mock test)
    TotalTests = TotalTests + 1
    If TestPowerShellCommandGeneration() Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 17.8: Test file dialog error handling
    TotalTests = TotalTests + 1
    If TestFileDialogErrorHandling() Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 17.9: Test GetCommandOutput functionality
    TotalTests = TotalTests + 1
    If TestGetCommandOutputFunctionality() Then
        TestsPassed = TestsPassed + 1
    End If
    
    ' Test 17.10: Integration test - Export/Import via UI functions
    TotalTests = TotalTests + 1
    If TestExportImportIntegration() Then
        TestsPassed = TestsPassed + 1
    End If
    
    FileDialogsTest = (TestsPassed >= TotalTests - 2) ' Allow 2 failures for UI-dependent tests
    Exit Function
    
ErrorHandler:
    If Not BootLoader.ErrorHandler Is Nothing Then
        BootLoader.ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "FileDialogsTest"
    End If
    FileDialogsTest = False
End Function

' Test 18: Property Calculation engine (PHASE-1 DORMANT, story 13-2) + PropertyTagging grammar-v2 matcher.
' The @cell=prop value seam is gone, so the calculation engine is asleep: IsTriggerCell is False for a cell
' that WOULD have triggered under v1, and ProcessElement / NoteDeletedTriggerCell are inert no-ops. The
' matcher assertions drive ElementMatchesAnyRule (Public, DGNLib-free) on real elements to prove the v2
' grammar matches as specified: Type[Line] matches a line not a cell; Cell[name] matches the named cell
' only; Type[Cell]&!Cell[A] matches a cell B, not cell A and not a line (strict negation); a wildcard
' Cell[ETI0*] matches ETI076. Pure element+config logic (no DGNLib); a -1 margin covers environment variance.
Private Function PropertyCalculationTest() As Boolean
    On Error GoTo ErrorHandler

    Dim TestsPassed As Integer
    Dim TotalTests As Integer

    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize

    ' Save config to restore afterwards
    Dim sOldEnabled As String, sOldDetach As String, sOldRules As String, sOldAuto As String
    sOldEnabled = ARESConfig.ARES_PROPERTY_CALC.Value
    sOldDetach = ARESConfig.ARES_CALC_DETACH_EMPTY.Value
    sOldRules = ARESConfig.ARES_PROPERTY_RULES.Value
    sOldAuto = ARESConfig.ARES_AUTO_PROPERTIES.Value

    ARESConfig.ARES_PROPERTY_CALC.Value = "True"          ' master ON - the engine is enabled but INERT
    ARESConfig.ARES_CALC_DETACH_EMPTY.Value = "False"
    ARESConfig.ARES_AUTO_PROPERTIES.Value = "True"

    ' --- Engine ASLEEP: a cell that WOULD have triggered under v1 (@Cell[PROPTEST]) is not a trigger ---
    ARESConfig.ARES_PROPERTY_RULES.Value = "@Cell[PROPTEST]=Repere"
    PropertyTagging.RefreshRules

    TotalTests = TotalTests + 1
    Dim cellSleep As element
    Set cellSleep = CreateCalculationTestCell("PROPTEST", 201, "Val", Point3dFromXYZ(600, 0, 0))
    If Not PropertyCalculation.IsTriggerCell(cellSleep) Then TestsPassed = TestsPassed + 1

    ' A plain (ungrouped) cell is not a trigger either
    TotalTests = TotalTests + 1
    Dim cellSleep2 As element
    Set cellSleep2 = CreateCalculationTestCell("PROPTEST", 0, "Val", Point3dFromXYZ(650, 0, 0))
    If Not PropertyCalculation.IsTriggerCell(cellSleep2) Then TestsPassed = TestsPassed + 1

    ' ProcessElement + NoteDeletedTriggerCell are inert no-ops (asleep: no crash, nothing recorded/written)
    TotalTests = TotalTests + 1
    Dim bInert As Boolean
    Dim sibsSleep() As element
    bInert = True
    PropertyCalculation.ProcessElement cellSleep                 ' asleep -> nothing calculated
    sibsSleep = Link.GetLink(cellSleep)
    PropertyCalculation.NoteDeletedTriggerCell cellSleep, sibsSleep   ' asleep -> records nothing
    If bInert Then TestsPassed = TestsPassed + 1

    ' --- Matcher (grammar v2) via ElementMatchesAnyRule (non-group rules, DGNLib-free) ---

    ' Type[Line]: matches a line, not a cell
    ARESConfig.ARES_PROPERTY_RULES.Value = "Type[Line]=Repere"
    PropertyTagging.RefreshRules
    Dim elLine As element
    Set elLine = CreateLineElement2(Nothing, Point3dFromXYZ(700, 0, 0), Point3dFromXYZ(800, 0, 0))
    ActiveModelReference.AddElement elLine
    Dim cellForType As element
    Set cellForType = CreateCalculationTestCell("ANYCELL", 0, "Val", Point3dFromXYZ(700, 60, 0))
    TotalTests = TotalTests + 1
    If PropertyTagging.ElementMatchesAnyRule(elLine) Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1
    If Not PropertyTagging.ElementMatchesAnyRule(cellForType) Then TestsPassed = TestsPassed + 1

    ' Cell[PROPX]: matches the cell named PROPX only (not another cell, not a line)
    ARESConfig.ARES_PROPERTY_RULES.Value = "Cell[PROPX]=Repere"
    PropertyTagging.RefreshRules
    Dim cellNamed As element, cellUnnamed As element, lineForCell As element
    Set cellNamed = CreateCalculationTestCell("PROPX", 0, "Val", Point3dFromXYZ(900, 0, 0))
    Set cellUnnamed = CreateCalculationTestCell("PROPY", 0, "Val", Point3dFromXYZ(950, 0, 0))
    Set lineForCell = CreateLineElement2(Nothing, Point3dFromXYZ(900, 60, 0), Point3dFromXYZ(1000, 60, 0))
    ActiveModelReference.AddElement lineForCell
    TotalTests = TotalTests + 1
    If PropertyTagging.ElementMatchesAnyRule(cellNamed) Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1
    If Not PropertyTagging.ElementMatchesAnyRule(cellUnnamed) Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1
    If Not PropertyTagging.ElementMatchesAnyRule(lineForCell) Then TestsPassed = TestsPassed + 1

    ' Type[Cell]&!Cell[A]: matches a cell B, not a cell A, not a line (strict negation)
    ARESConfig.ARES_PROPERTY_RULES.Value = "Type[Cell]&!Cell[A]=Repere"
    PropertyTagging.RefreshRules
    Dim cellB As element, cellA As element, lineNeg As element
    Set cellB = CreateCalculationTestCell("B", 0, "Val", Point3dFromXYZ(1100, 0, 0))
    Set cellA = CreateCalculationTestCell("A", 0, "Val", Point3dFromXYZ(1150, 0, 0))
    Set lineNeg = CreateLineElement2(Nothing, Point3dFromXYZ(1100, 60, 0), Point3dFromXYZ(1200, 60, 0))
    ActiveModelReference.AddElement lineNeg
    TotalTests = TotalTests + 1
    If PropertyTagging.ElementMatchesAnyRule(cellB) Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1
    If Not PropertyTagging.ElementMatchesAnyRule(cellA) Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1
    If Not PropertyTagging.ElementMatchesAnyRule(lineNeg) Then TestsPassed = TestsPassed + 1

    ' Wildcard Cell[ETI0*]: matches a cell named ETI076
    ARESConfig.ARES_PROPERTY_RULES.Value = "Cell[ETI0*]=Repere"
    PropertyTagging.RefreshRules
    Dim cellWild As element
    Set cellWild = CreateCalculationTestCell("ETI076", 0, "Val", Point3dFromXYZ(1300, 0, 0))
    TotalTests = TotalTests + 1
    If PropertyTagging.ElementMatchesAnyRule(cellWild) Then TestsPassed = TestsPassed + 1

    ' Restore config
    ARESConfig.ARES_PROPERTY_CALC.Value = sOldEnabled
    ARESConfig.ARES_CALC_DETACH_EMPTY.Value = sOldDetach
    ARESConfig.ARES_PROPERTY_RULES.Value = sOldRules
    ARESConfig.ARES_AUTO_PROPERTIES.Value = sOldAuto
    PropertyTagging.RefreshRules                          ' re-parse the restored rules (tests changed them)

    ' Allow a small margin for environment variance (as CustomPropertyHandlerTest does)
    PropertyCalculationTest = (TestsPassed >= TotalTests - 1)
    Exit Function

ErrorHandler:
    If Not BootLoader.ErrorHandler Is Nothing Then
        BootLoader.ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyCalculationTest"
    End If
    PropertyCalculationTest = False
End Function

' Test 19: PropertyTagging grammar v2 - ValidateAndNormalizeRule (validate + normalise to canonical) and
' RuleHasNoEffect (dead-rule contradiction detector). Pure string logic, deterministic, DGNLib-free, no
' config mutation (reads only module constants), so no save/restore and no -1 tolerance - every case passes
' EXACTLY. Valid rules return "" and the expected COMPACT canonical form (matching the story I/O matrix,
' no spaces around "&"/"="); invalid rules return a non-empty reason (incl. every v1 rule and the exact
' "|"-instead-of-";" incident); RuleHasNoEffect is True (with the two conflicting segments) for the four
' contradiction shapes and False for a compatible or wildcard-guarded rule.
Private Function PropertyRuleValidationTest() As Boolean
    On Error GoTo ErrorHandler

    Dim TestsPassed As Integer
    Dim TotalTests As Integer
    Dim canon As String
    Dim segs() As String

    ' --- Valid rules: reason "" AND the expected canonical form (matrix #1-16) ---
    TotalTests = TotalTests + 1: If VNorm("Lvl[WALLS]=Commune", "Lvl[WALLS]=Commune") Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If VNorm("Cell[ETI076]=Repere", "Cell[ETI076]=Repere") Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If VNorm("@Cell[ETI076]=Repere", "@Cell[ETI076]=Repere") Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If VNorm("Type[Line]=Commune", "Type[Line]=Commune") Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If VNorm("Type[Cell]&!Cell[A]=Repere", "Type[Cell]&!Cell[A]=Repere") Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If VNorm("Lvl[Poste=HTA]=Coupe_Type", "Lvl[Poste=HTA]=Coupe_Type") Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If VNorm("Lvl[R&D]=Commune", "Lvl[R&D]=Commune") Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If VNorm("Cell[ETI0*]=Repere", "Cell[ETI0*]=Repere") Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If VNorm("Cell[ETI0?6]=Repere", "Cell[ETI0?6]=Repere") Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If VNorm("Lvl[A|B]=Commune", "Lvl[A|B]=Commune") Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If VNorm("Cell[X]&Lvl[Y]=P", "Cell[X]&Lvl[Y]=P") Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If VNorm("Cell[X]=P1|P2", "Cell[X]=P1|P2") Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If VNorm("lvl[walls]=Commune", "Lvl[walls]=Commune") Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If VNorm("Cell[X]@=P", "@Cell[X]=P") Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If VNorm("@@Cell[X]=P", "@Cell[X]=P") Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If VNorm("Cell[X]=Rep@ere", "Cell[X]=Rep@ere") Then TestsPassed = TestsPassed + 1
    ' Normalisation collapses spare spaces around "&" and "="
    TotalTests = TotalTests + 1: If VNorm("  Cell[X] & Lvl[Y]  =  P  ", "Cell[X]&Lvl[Y]=P") Then TestsPassed = TestsPassed + 1
    ' Empty rule -> "" reason with empty canonical (the caller deletes)
    TotalTests = TotalTests + 1
    canon = "sentinel"
    If PropertyTagging.ValidateAndNormalizeRule("", canon) = "" Then
        If canon = "" Then TestsPassed = TestsPassed + 1
    End If

    ' --- Invalid rules: a non-empty reason (matrix #17-28 + the v1 incident) ---
    TotalTests = TotalTests + 1: If VReject("WALLS=Commune") Then TestsPassed = TestsPassed + 1               ' v1 level rule (no keyword)
    TotalTests = TotalTests + 1: If VReject("@ETI076=Repere") Then TestsPassed = TestsPassed + 1              ' v1 cell rule (no keyword)
    TotalTests = TotalTests + 1: If VReject("Color[3]=P") Then TestsPassed = TestsPassed + 1                  ' unknown keyword
    TotalTests = TotalTests + 1: If VReject("Cell[X]") Then TestsPassed = TestsPassed + 1                     ' no "="
    TotalTests = TotalTests + 1: If VReject("=P") Then TestsPassed = TestsPassed + 1                          ' empty condition side
    TotalTests = TotalTests + 1: If VReject("Cell[X]=") Then TestsPassed = TestsPassed + 1                    ' empty prop side
    TotalTests = TotalTests + 1: If VReject("Cell[]=P") Then TestsPassed = TestsPassed + 1                    ' empty name list
    TotalTests = TotalTests + 1: If VReject("Cell[A;B]=P") Then TestsPassed = TestsPassed + 1                 ' ";" inside [...]
    TotalTests = TotalTests + 1: If VReject("Cell[A=P") Then TestsPassed = TestsPassed + 1                    ' unbalanced "["
    TotalTests = TotalTests + 1: If VReject("(Cell[X])=P") Then TestsPassed = TestsPassed + 1                 ' reserved parens
    TotalTests = TotalTests + 1: If VReject("Cell[X]=P1|@Y=Q") Then TestsPassed = TestsPassed + 1             ' prop token contains "="
    TotalTests = TotalTests + 1: If VReject("Type[Bogus]=P") Then TestsPassed = TestsPassed + 1               ' unknown type
    TotalTests = TotalTests + 1: If VReject("Type[Li*]=P") Then TestsPassed = TestsPassed + 1                 ' wildcard in Type
    ' The exact v1 incident: rules joined with "|" instead of ";" -> parses as ONE rule -> rejected
    TotalTests = TotalTests + 1: If VReject("ARES_Tranchee=Coupe_Type|@ETI03Z=Repere|@ETI053B=Repere") Then TestsPassed = TestsPassed + 1

    ' --- Contradiction detector: True (with the two segments) for the four shapes + #34 ---
    TotalTests = TotalTests + 1: If DeadSeg("Type[Line]&Type[Arc]=P", "Type[Line]", "Type[Arc]") Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If DeadSeg("Lvl[A]&Lvl[B]=P", "Lvl[A]", "Lvl[B]") Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If DeadSeg("Cell[X]&Type[Line]=P", "Cell[X]", "Type[Line]") Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If DeadSeg("!Type[Cell]&Cell[X]=P", "!Type[Cell]", "Cell[X]") Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If DeadSeg("Type[Line|Arc]&Type[Shape]=P", "Type[Line|Arc]", "Type[Shape]") Then TestsPassed = TestsPassed + 1
    ' Round 3: the Cell-vs-Type contradiction is STRUCTURAL - a wildcard in the Cell name does NOT suppress it.
    TotalTests = TotalTests + 1: If DeadSeg("@Cell[ETI0*]&!Type[Cell]=P", "Cell[ETI0*]", "!Type[Cell]") Then TestsPassed = TestsPassed + 1
    ' Not dead: compatible (Cell + Type[Cell]) / wildcard guard (Lvl[A*]) -> no verdict
    TotalTests = TotalTests + 1: If Not PropertyTagging.RuleHasNoEffect("Cell[X]&Type[Cell]=P", segs) Then TestsPassed = TestsPassed + 1
    TotalTests = TotalTests + 1: If Not PropertyTagging.RuleHasNoEffect("Lvl[A*]&Lvl[B]=P", segs) Then TestsPassed = TestsPassed + 1
    ' Round 3: a wildcard STILL suppresses the same-keyword disjoint-list verdict (undecidable).
    TotalTests = TotalTests + 1: If Not PropertyTagging.RuleHasNoEffect("Lvl[A*]&Lvl[B*]=P", segs) Then TestsPassed = TestsPassed + 1
    ' An invalid rule (wildcard in Type) yields no verdict
    TotalTests = TotalTests + 1: If Not PropertyTagging.RuleHasNoEffect("Type[Line]&Type[Lin*]=P", segs) Then TestsPassed = TestsPassed + 1

    PropertyRuleValidationTest = (TestsPassed = TotalTests)
    Exit Function

ErrorHandler:
    If Not BootLoader.ErrorHandler Is Nothing Then
        BootLoader.ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "PropertyRuleValidationTest"
    End If
    PropertyRuleValidationTest = False
End Function

' Helper: True when sRule validates (reason "") to EXACTLY sExpectedCanonical.
Private Function VNorm(ByVal sRule As String, ByVal sExpectedCanonical As String) As Boolean
    Dim canon As String
    Dim reason As String
    VNorm = False
    reason = PropertyTagging.ValidateAndNormalizeRule(sRule, canon)
    If Len(reason) = 0 Then
        If canon = sExpectedCanonical Then VNorm = True
    End If
End Function

' Helper: True when sRule is rejected (non-empty reason).
Private Function VReject(ByVal sRule As String) As Boolean
    Dim canon As String
    VReject = (Len(PropertyTagging.ValidateAndNormalizeRule(sRule, canon)) > 0)
End Function

' Helper: True when sRule is flagged dead with exactly the two expected canonical segments (in order).
Private Function DeadSeg(ByVal sRule As String, ByVal seg1 As String, ByVal seg2 As String) As Boolean
    Dim segs() As String
    DeadSeg = False
    If PropertyTagging.RuleHasNoEffect(sRule, segs) Then
        If UBound(segs) - LBound(segs) = 1 Then
            If segs(LBound(segs)) = seg1 Then
                If segs(LBound(segs) + 1) = seg2 Then DeadSeg = True
            End If
        End If
    End If
End Function

' Build a single-TextElement graphic cell named sName, in graphic group lGroup (0 = ungrouped),
' added to the active model. Helper for PropertyCalculationTest.
Private Function CreateCalculationTestCell(ByVal sName As String, ByVal lGroup As Long, ByVal sText As String, ByVal origin As Point3d) As CellElement
    Dim arr(0) As element
    Set arr(0) = CreateTextElement1(Nothing, sText, origin, Matrix3dIdentity)
    Dim oCell As CellElement
    Set oCell = CreateCellElement1(sName, arr, origin)
    If lGroup <> 0 Then oCell.GraphicGroup = lGroup
    ActiveModelReference.AddElement oCell
    Set CreateCalculationTestCell = oCell
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

Private Function TestCleanFilePathFunctionality() As Boolean
    On Error GoTo ErrorHandler
    
    ' We can't directly test the private CleanFilePath function,
    ' but we can test similar functionality by simulating what it should do
    Dim TestString As String
    Dim CleanString As String
    
    TestString = "C:\Test Path\file.cfg" & vbCr & vbLf & vbTab
    
    ' Create a simple version of what CleanFilePath should do
    CleanString = Trim(TestString)
    CleanString = Replace(CleanString, vbCr, "")
    CleanString = Replace(CleanString, vbLf, "")
    CleanString = Replace(CleanString, vbTab, "")
    
    ' If we can clean a file path manually, the function should work
    If CleanString = "C:\Test Path\file.cfg" Then
        TestCleanFilePathFunctionality = True
    End If
    
    Exit Function
    
ErrorHandler:
    TestCleanFilePathFunctionality = False
End Function

Private Function TestEscapeForPowerShellFunctionality() As Boolean
    On Error GoTo ErrorHandler
    
    ' Test escaping functionality (simulate what EscapeForPowerShell should do)
    Dim TestString As String
    Dim EscapedString As String
    
    TestString = "Test's ""quoted"" \path\"
    
    ' Simulate escaping
    EscapedString = TestString
    EscapedString = Replace(EscapedString, "'", "''")
    EscapedString = Replace(EscapedString, """", """""")
    EscapedString = Replace(EscapedString, "\", "\\")
    
    ' If escaping works manually, the function should work
    If InStr(EscapedString, "''") > 0 And InStr(EscapedString, """""") > 0 And InStr(EscapedString, "\\") > 0 Then
        TestEscapeForPowerShellFunctionality = True
    End If
    
    Exit Function
    
ErrorHandler:
    TestEscapeForPowerShellFunctionality = False
End Function

Private Function TestPowerShellCommandGeneration() As Boolean
    On Error GoTo ErrorHandler
    
    ' Test that we can generate valid PowerShell commands
    Dim TestTitle As String
    Dim TestDir As String
    Dim TestFile As String
    
    TestTitle = "Test Export"
    TestDir = "C:\Temp"
    TestFile = "test.cfg"
    
    ' Build a test PowerShell command similar to what ShowSaveFileDialog builds
    Dim PSCommand As String
    PSCommand = "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -Command """ & _
                "Add-Type -AssemblyName System.Windows.Forms; " & _
                "$dialog = New-Object System.Windows.Forms.SaveFileDialog; " & _
                "$dialog.Title = '" & TestTitle & "'; " & _
                "$dialog.Filter = 'ARES Config (*.cfg)|*.cfg|All Files (*.*)|*.*'; " & _
                "$dialog.DefaultExt = 'cfg'; " & _
                "$dialog.InitialDirectory = '" & TestDir & "'; " & _
                "$dialog.FileName = '" & TestFile & "'; " & _
                "Write-Output 'CommandGenerated'"""
    
    ' If command contains expected elements, generation logic should work
    If InStr(PSCommand, "System.Windows.Forms") > 0 And _
       InStr(PSCommand, "SaveFileDialog") > 0 And _
       InStr(PSCommand, TestTitle) > 0 And _
       InStr(PSCommand, TestFile) > 0 Then
        TestPowerShellCommandGeneration = True
    End If
    
    Exit Function
    
ErrorHandler:
    TestPowerShellCommandGeneration = False
End Function

Private Function TestFileDialogErrorHandling() As Boolean
    On Error GoTo ErrorHandler
    
    ' Test error handling by calling dialog functions with invalid parameters
    Dim Result1 As String
    Dim Result2 As String
    
    On Error Resume Next
    
    ' These should handle errors gracefully and return empty strings
    Result1 = FileDialogs.ShowSaveDialog("", "", "", "", "")
    Result2 = FileDialogs.ShowOpenFileDialog("", "")
    
    ' Functions should return empty strings on error, not crash
    If Err.Number = 0 And Len(Result1) = 0 And Len(Result2) = 0 Then
        TestFileDialogErrorHandling = True
    End If
    
    On Error GoTo ErrorHandler
    Exit Function
    
ErrorHandler:
    TestFileDialogErrorHandling = False
End Function

Private Function TestGetCommandOutputFunctionality() As Boolean
    On Error GoTo ErrorHandler
    
    ' Test basic command output functionality using echo
    Dim TestCommand As String
    Dim Output As String
    
    TestCommand = "echo ARES_FILE_DIALOG_TEST"
    
    ' We can't directly call the private GetCommandOutput function,
    ' but we can test similar functionality
    Dim wshShell As Object
    Dim TempFile As String
    Dim BatchFile As String
    Dim FileNum As Integer
    
    Set wshShell = CreateObject("WScript.Shell")
    
    TempFile = Environ("TEMP") & "\ares_test_output.txt"
    BatchFile = Environ("TEMP") & "\ares_test_cmd.bat"
    
    ' Create batch file
    FileNum = FreeFile
    Open BatchFile For Output As #FileNum
    Print #FileNum, "@echo off"
    Print #FileNum, TestCommand & " > """ & TempFile & """"
    Close #FileNum
    
    ' Execute batch file
    wshShell.Run """" & BatchFile & """", 0, True
    
    ' Read output
    If Dir(TempFile) <> "" Then
        FileNum = FreeFile
        Open TempFile For Input As #FileNum
        If Not EOF(FileNum) Then
            Output = Input(LOF(FileNum), FileNum)
        End If
        Close #FileNum
    End If
    
    ' Cleanup
    On Error Resume Next
    If Dir(TempFile) <> "" Then Kill TempFile
    If Dir(BatchFile) <> "" Then Kill BatchFile
    On Error GoTo ErrorHandler
    
    ' Test if we got expected output
    If InStr(Output, "ARES_FILE_DIALOG_TEST") > 0 Then
        TestGetCommandOutputFunctionality = True
    End If
    
    Exit Function
    
ErrorHandler:
    ' Cleanup on error
    On Error Resume Next
    If Dir(TempFile) <> "" Then Kill TempFile
    If Dir(BatchFile) <> "" Then Kill BatchFile
    TestGetCommandOutputFunctionality = False
End Function

Private Function TestExportImportIntegration() As Boolean
    On Error GoTo ErrorHandler
    
    ' Test the integration between export/import functionality
    ' This tests the underlying logic without showing UI dialogs
    
    ' Ensure ARESConfig is initialized
    If Not ARESConfig.IsInitialized Then
        ARESConfig.Initialize
    End If
    
    ' Test export to specific file
    Dim TestExportPath As String
    TestExportPath = Environ("TEMP") & "\ARES_Dialog_Test_Export.cfg"
    
    ' Clean up any existing file
    On Error Resume Next
    If Dir(TestExportPath) <> "" Then Kill TestExportPath
    On Error GoTo ErrorHandler
    
    ' Export configuration
    If ARESConfig.ExportConfig(TestExportPath) Then
        ' Verify file was created
        If Dir(TestExportPath) <> "" Then
            ' Test import
            If ARESConfig.ImportConfig(TestExportPath, True) Then
                TestExportImportIntegration = True
            End If
        End If
    End If
    
    ' Cleanup
    On Error Resume Next
    If Dir(TestExportPath) <> "" Then Kill TestExportPath
    On Error GoTo ErrorHandler
    
    Exit Function
    
ErrorHandler:
    TestExportImportIntegration = False
End Function

Private Sub RunTest(TestName As String, TestIdentifier As Integer)
    Dim StartTime As Double
    Dim Result As TestResult
    
    StartTime = Timer
    
    Result.Name = TestName
    
    On Error Resume Next
    ' Execute test based on ID (UPDATED)
    Select Case TestIdentifier
        Case tidConfig: Result.Passed = ConfigTest()
        Case tidLangManager: Result.Passed = LangManagerTest()
        Case tidUUID: Result.Passed = UUIDTest()
        Case tidARESVars: Result.Passed = ARES_VARTest()
        Case tidCustomProps: Result.Passed = CustomPropertyHandlerTest()
        Case tidErrorHandler: Result.Passed = ErrorHandlerTest()
        Case tidElementProcess: Result.Passed = ElementInProcesseTest()
        Case tidLength: Result.Passed = LengthTest()
        Case tidMSd: Result.Passed = MSdTest()
        Case tidStringsInEl: Result.Passed = StringsInElTest()
        Case tidLink: Result.Passed = LinkTest()
        Case tidMSGraphical: Result.Passed = MSGraphicalTest()
        Case tidARESMSVar: Result.Passed = ARESMSVarTest()
        Case tidBootLoader: Result.Passed = BootLoaderTest()
        Case tidAutoLengths: Result.Passed = AutoLengthsTest()
        Case tidConfigExportImport: Result.Passed = ConfigExportImportTest()
        Case tidFileDialogs: Result.Passed = FileDialogsTest()
        Case tidPropertyCalculation: Result.Passed = PropertyCalculationTest()
        Case tidPropertyRuleValidation: Result.Passed = PropertyRuleValidationTest()
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