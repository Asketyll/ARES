' Module: Command
' Description: Liste all command
' License: This project is licensed under the AGPL-3.0.
' Dependencies: AutoLengths, BootLoader, LangManager
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
