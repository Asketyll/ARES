Sub MasterTest()

    ConfigTest
    LangManagerTest
    UUIDTest
    ARES_VARTest
    CustomPropertyHandlerTest
End Sub
Sub ConfigTest()
    ShowStatus Config.SetVar("ARES_Unit_testing", "I'm a test unit variable")
    ShowStatus Config.GetVar("ARES_Unit_testing")
    ShowStatus Config.RemoveValue("ARES_Unit_testing")
End Sub
Sub LangManagerTest()
    LangManager.English
    ShowStatus LangManager.GetTranslation("UnitTesting")
End Sub
Sub UUIDTest()
    ShowStatus uuid.GenerateV1
End Sub
Sub ARES_VARTest()
    ShowStatus ARES_VAR.InitMSVars
    ShowStatus ARES_VAR.RemoveMSVar("ARES_UnitTesting", False)
    ShowStatus ARES_VAR.ResetMSVar("ARES_UnitTesting")
End Sub
Sub CustomPropertyHandlerTest()
    CustomPropertyHandler
End Sub
