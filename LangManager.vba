' Module: LangManager
' Description: This module manages translations for different languages in GUI.

' Dependencies: Config, ARES_VAR

Option Explicit

Public Enum SupportedLanguages
    English = 1
    French = 2
    ' Add more languages as needed
End Enum

Private mTranslations As Object

Sub InitializeTranslations()
    Set mTranslations = CreateObject("Scripting.Dictionary")

    ' Add English translations
    mTranslations.Add "EN_VarResetSuccess", "Reset to default value: {0}"
    mTranslations.Add "EN_VarResetError", "Unable to reset the variable."
    mTranslations.Add "EN_VarRemoveConfirm", "Do you really want to remove the variable {0} ?"
    mTranslations.Add "EN_VarRemoveSuccess", "Removed."
    mTranslations.Add "EN_VarRemoveError", "Unable to remove the variable."
    mTranslations.Add "EN_VarKeyNotFound", "Key not found in the collection: {0}"
    mTranslations.Add "EN_VarInvalidArgument", "Invalid argument type."
    mTranslations.Add "EN_VarInitializeMSVarfailed", "ARES Config with MS Vars failed."
    mTranslations.Add "EN_VarKeyNotInCollection", "The variable: {0} is not known."
    mTranslations.Add "EN_VarsRemoveConfirm", "Do you really want to remove all variables ?"
    mTranslations.Add "EN_BootUserLangInit", "User language initialized."
    mTranslations.Add "EN_BootMSVarsInit", "Variable management initialized."
    mTranslations.Add "EN_BootMSVarsMissing", "Variable management is missing."
    mTranslations.Add "EN_BootFail", "Error in automatic loading of VBA."
    mTranslations.Add "EN_LangFail", "Translation not found for key: "
    mTranslations.Add "EN_LengthRoundError", "Rounding value unauthorized: "
    mTranslations.Add "EN_LengthElementTypeNotSupportedByInterface", "The element: {0} is an element of type: {1}, it is not supported by the GetElementLength interface."
    mTranslations.Add "EN_AutoLengthsUpdateError", "An error occurred while updating lengths."
    mTranslations.Add "EN_DGNOpenCloseEventsInitialized", "Track events element initialized."
    mTranslations.Add "EN_AutoLengthsGUIInvalidSelectedElement", "The selected item is invalid."
    mTranslations.Add "EN_AutoLengthsGUICaption", "Select:"
    
    ' Add French translations
    mTranslations.Add "FR_VarResetSuccess", "Réinitialisé à la valeur par défaut: {0}"
    mTranslations.Add "FR_VarResetError", "Impossible de réinitialiser la variable."
    mTranslations.Add "FR_VarRemoveConfirm", "Voulez-vous vraiment supprimer la variable {0}?"
    mTranslations.Add "FR_VarRemoveSuccess", "Supprimé."
    mTranslations.Add "FR_VarRemoveError", "Impossible de supprimer la variable."
    mTranslations.Add "FR_VarKeyNotFound", "Clé introuvable dans la collection: {0}"
    mTranslations.Add "FR_VarInvalidArgument", "Type d'argument non valide."
    mTranslations.Add "FR_VarInitializeMSVarfailed", "ARES Config avec MS Vars à échoué."
    mTranslations.Add "FR_VarKeyNotInCollection", "La variable: {0} n'est pas reconnue."
    mTranslations.Add "FR_VarsRemoveConfirm", "Voulez-vous vraiment supprimer toutes les variables ?"
    mTranslations.Add "FR_BootUserLangInit", "Langage utilisateur initialisé."
    mTranslations.Add "FR_BootMSVarsInit", "Gestion des variables initialisées."
    mTranslations.Add "FR_BootMSVarsMissing", "Gestion des variables manquante."
    mTranslations.Add "FR_BootFail", "Erreur lors du chargement automatique de VBA."
    mTranslations.Add "FR_LangFail", "Traduction introuvable pour la clé: "
    mTranslations.Add "FR_LengthRoundError", "Valeur d'arrondi interdit: "
    mTranslations.Add "FR_LengthElementTypeNotSupportedByInterface", "L'élément: {0} est un élément de type: {1}, il n'est pas géré par l'interface GetElementLength."
    mTranslations.Add "FR_AutoLengthsUpdateError", "Une erreur s'est produite lors de la mise à jour des longueurs."
    mTranslations.Add "FR_DGNOpenCloseEventsInitialized", "Evénements de suivi d'objet initialisé."
    mTranslations.Add "FR_AutoLengthsGUIInvalidSelectedElement", "L'élément sélectionné n'est pas valide."
    mTranslations.Add "FR_AutoLengthsGUICaption", "Sélectionner:"
End Sub

Public Function GetTranslation(key As String, ParamArray params() As Variant) As String
    Dim baseKey As String
    Dim formattedMessage As String
    Dim i As Integer
    
    baseKey = UCase(Left(GetUserLanguage, 2)) & "_" & key
    
    ' Retrieve the message based on the constructed key
    If mTranslations.Exists(baseKey) Then
        GetTranslation = mTranslations(baseKey)
    Else
        ' Default to English if translation not found
        baseKey = "EN_" & key
        If mTranslations.Exists(baseKey) Then
            GetTranslation = mTranslations(baseKey)
        Else
            GetTranslation = GetTranslation("LangFail") & key
            Exit Function
        End If
    End If

    ' Format the message with parameters if any are provided
    If UBound(params) - LBound(params) >= 0 Then
        formattedMessage = GetTranslation
        For i = LBound(params) To UBound(params)
            formattedMessage = Replace(formattedMessage, "{" & i & "}", params(i))
        Next i
        GetTranslation = formattedMessage
    End If
End Function

Public Function GetUserLanguage() As String
    GetUserLanguage = Config.GetVar("CONNECTUSER_LANGUAGE")
    If GetUserLanguage = "" Or GetUserLanguage = ARES_VAR.ARES_NAVD Then
        If ARES_VAR.ARES_LANGUAGE Is Nothing Then
            ARES_VAR.InitMSVars
        End If
        GetUserLanguage = ARES_VAR.ARES_LANGUAGE.Value
    End If
End Function
