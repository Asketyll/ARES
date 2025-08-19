' Module: Command
' Description: Liste all command
' License: This project is licensed under the AGPL-3.0.
' Dependencies: AutoLengths, BootLoader, LangManager, ARESConfigClass
Option Explicit

' Sub to call CommandState for manual update length in string
Sub ForceUpdateLength()
    CommandState.StartLocate New AutoLengths
End Sub

' Sub to reset all ARES var in MS
Sub ResetARESVariables()
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If
    If ARESConfig.ResetAllConfigVars() Then
        If Not LangManager.IsInit Then
            LangManager.InitializeTranslations
        End If
        ShowStatus GetTranslation("VarResetAllSuccess")
    Else
        ShowStatus GetTranslation("VarResetAllFailed")
    End If
End Sub

' Sub to remove all ARES var in MS
Sub RemoveARESVariables()
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If
    If ARESConfig.RemoveAllConfigVars() Then
        If Not LangManager.IsInit Then
            LangManager.InitializeTranslations
        End If
        ShowStatus GetTranslation("VarRemoveSuccess")
    Else
        ShowStatus GetTranslation("VarRemoveError")
    End If
End Sub

' Sub to call GUI Options of AutoLenghts
Sub EditAutoLenghtsOptions()
    If BootLoader.ARESConfig Is Nothing Or Not ARESConfig.IsInitialized Then
        Set BootLoader.ARESConfig = New ARESConfigClass
        ARESConfig.Initialize
    End If
    If Not LangManager.IsInit Then
        LangManager.InitializeTranslations
    End If
    Dim frm As New AutoLenghts_GUI_Options
    frm.Show vbModeless
End Sub

' Sub to set language to English
Sub English()
    If Config.SetVar(ARESConfig.ARES_LANGUAGE.key, "English") Then
        ShowStatus ARESConfig.ARES_LANGUAGE.key & " set to English, please restart."
    Else
        ShowStatus "Imposible to set " & ARESConfig.ARES_LANGUAGE.key & ", please try manualy."
    End If
End Sub

' Sub to set language to French
Sub Français()
    If Config.SetVar(ARESConfig.ARES_LANGUAGE.key, "Français") Then
        ShowStatus ARESConfig.ARES_LANGUAGE.key & " défini à Français, veuillez redémarrer."
    Else
        ShowStatus "Impossible de définir " & ARESConfig.ARES_LANGUAGE.key & ", veuillez essayer manuellement."
    End If
End Sub
