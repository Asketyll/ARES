' Module: LangManager
' Description: This module manages translations for different languages in GUI.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: Config, ARES_VAR

Option Explicit

Private mSupportedLanguages As Collection
Private mTranslations As Object
Private mUserLanguage As String
Public IsInit As Boolean

' Initialize translations and supported languages
Sub InitializeTranslations()
    On Error GoTo ErrorHandler

    Set mSupportedLanguages = New Collection
    Set mTranslations = CreateObject("Scripting.Dictionary")
    
    ' Add supported languages to the collection
    mSupportedLanguages.Add "English"
    mSupportedLanguages.Add "Français"

    ' Initialize user language
    mUserLanguage = GetUserLanguage()
    
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
    mTranslations.Add "EN_VarsRemoveConfirm", "Do you really want to remove all variables ? This action is irreversible."
    mTranslations.Add "EN_BootUserLangInit", "User language initialized."
    mTranslations.Add "EN_BootMSVarsInit", "Variable management initialized."
    mTranslations.Add "EN_BootMSVarsMissing", "Variable management is missing."
    mTranslations.Add "EN_BootFail", "Error in automatic loading of VBA."
    mTranslations.Add "EN_LangFail", "Translation not found for key: "
    mTranslations.Add "EN_LengthRoundError", "Rounding value unauthorized: "
    mTranslations.Add "EN_LengthElementTypeNotSupportedByInterface", "The element: {0} is an element of type: {1}, it is not supported by the GetElementLength interface."
    mTranslations.Add "EN_DGNOpenCloseEventsInitialized", "Track events element initialized."
    mTranslations.Add "EN_DGNOpenCloseInitError", "Error initializing DGN Open/Close events: "
    mTranslations.Add "EN_AutoLengthsGUIInvalidSelectedElement", "The selected item is invalid."
    mTranslations.Add "EN_AutoLengthsGUICaption", "Select:"
    mTranslations.Add "EN_AutoLengthsInitError", "Error initializing AutoLengths: "
    mTranslations.Add "EN_AutoLengthsCalculationError", "Error calculating lengths: "
    mTranslations.Add "EN_AutoLengthsUpdateError", "An error occurred while updating lengths: "
    mTranslations.Add "EN_AutoLengthsShowFormError", "Error showing element selection form: "
    mTranslations.Add "EN_AutoLengthsSelectionError", "Error selecting element: "
    mTranslations.Add "EN_AutoLengthsSetTriggerError", "Error setting trigger: "
    mTranslations.Add "EN_AutoLengthsAddTriggerError", "Error adding trigger: "
    mTranslations.Add "EN_AutoLengthsResetTriggerError", "Error resetting trigger: "
    mTranslations.Add "EN_AutoLengthsNoValidElement", "No valid element selected"
    mTranslations.Add "EN_AutoLengthsSelectanelementC", "Select an element"
    mTranslations.Add "EN_AutoLengthsSelectanelementP", "Select a valid and linked element"
    
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
    mTranslations.Add "FR_VarsRemoveConfirm", "Voulez-vous vraiment supprimer toutes les variables ? Cette action est irréversible."
    mTranslations.Add "FR_BootUserLangInit", "Langage utilisateur initialisé."
    mTranslations.Add "FR_BootMSVarsInit", "Gestion des variables initialisées."
    mTranslations.Add "FR_BootMSVarsMissing", "Gestion des variables manquante."
    mTranslations.Add "FR_BootFail", "Erreur lors du chargement automatique de VBA."
    mTranslations.Add "FR_LangFail", "Traduction introuvable pour la clé: "
    mTranslations.Add "FR_LengthRoundError", "Valeur d'arrondi interdit: "
    mTranslations.Add "FR_LengthElementTypeNotSupportedByInterface", "L'élément: {0} est un élément de type: {1}, il n'est pas géré par l'interface GetElementLength."
    mTranslations.Add "FR_DGNOpenCloseEventsInitialized", "Evénements de suivi d'objet initialisé."
    mTranslations.Add "FR_DGNOpenCloseInitError", "Erreur lors de l'initialisation des événements d'ouverture/fermeture DGN: "
    mTranslations.Add "FR_AutoLengthsGUIInvalidSelectedElement", "L'élément sélectionné n'est pas valide."
    mTranslations.Add "FR_AutoLengthsGUICaption", "Sélectionner:"
    mTranslations.Add "FR_AutoLengthsInitError", "Erreur lors de l'initialisation d'AutoLengths: "
    mTranslations.Add "FR_AutoLengthsCalculationError", "Erreur lors du calcul des longueurs: "
    mTranslations.Add "FR_AutoLengthsUpdateError", "Une erreur s'est produite lors de la mise à jour des longueurs: "
    mTranslations.Add "FR_AutoLengthsShowFormError", "Erreur lors de l'affichage du formulaire de sélection d'élément: "
    mTranslations.Add "FR_AutoLengthsSelectionError", "Erreur lors de la sélection de l'élément: "
    mTranslations.Add "FR_AutoLengthsSetTriggerError", "Erreur lors de la définition du déclencheur: "
    mTranslations.Add "FR_AutoLengthsAddTriggerError", "Erreur lors de l'ajout du déclencheur: "
    mTranslations.Add "FR_AutoLengthsResetTriggerError", "Erreur lors de la réinitialisation du déclencheur: "
    mTranslations.Add "FR_AutoLengthsNoValidElement", "Aucun élément valide sélectionné"
    mTranslations.Add "FR_AutoLengthsSelectanelementC", "Sélectionner un élément"
    mTranslations.Add "FR_AutoLengthsSelectanelementP", "Sélectionner un élément valide et groupé"
    IsInit = True
    
    Exit Sub

ErrorHandler:
    IsInit = False
    MsgBox "An error occurred while initializing translations.", vbOKOnly
End Sub

' Function to get translation based on key and parameters
Public Function GetTranslation(key As String, ParamArray params() As Variant) As String
    On Error GoTo ErrorHandler

    Dim baseKey As String
    Dim formattedMessage As String
    Dim i As Integer
    
    baseKey = UCase(Left(mUserLanguage, 2)) & "_" & key
    
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

    Exit Function

ErrorHandler:
    GetTranslation = "Error retrieving translation for key: " & key
End Function

' Function to get the user language
Private Function GetUserLanguage() As String
    On Error GoTo ErrorHandler

    Dim Prompt As String
    Dim lang As Variant
    
    GetUserLanguage = Config.GetVar("CONNECTUSER_LANGUAGE")

    If GetUserLanguage = "" Or GetUserLanguage = ARES_VAR.ARES_NAVD Then
        If ARES_VAR.ARES_LANGUAGE Is Nothing Then
            ARES_VAR.InitMSVars
        End If

        If ARES_VAR.ARES_LANGUAGE.Value = "" Then
            Prompt = "Unable to retrieve your user language." & vbCrLf & _
            "We invite you to declare it in the MicroStation environment variable, key: ARES_Language" & vbCrLf & _
            "The supported languages are:"
            
            ' Loop through the enum values and append them to the message
            For Each lang In mSupportedLanguages
                Prompt = Prompt & vbCrLf & "- " & lang
            Next lang
            
            Prompt = Prompt & vbCrLf & vbCrLf & "You can use the keyin: '""macro vba run [ARES]English'"" or '""macro vba run [ARES]Français'"""
            'macro vba run Ares.SetLanguage
            MsgBox Prompt, vbOKOnly, "User language"
        Else
            GetUserLanguage = ARES_VAR.ARES_LANGUAGE.Value
        End If
    End If

    Exit Function

ErrorHandler:
    GetUserLanguage = "English" ' Default language in case of error
End Function

' Sub to set language to English
Sub English()
    If Config.SetVar(ARES.ARES_LANGUAGE.key, "English") Then
        ShowStatus ARES.ARES_LANGUAGE.key & " set to English, please restart."
    Else
        ShowStatus "Imposible to set " & ARES.ARES_LANGUAGE.key & ", please try manualy."
    End If
End Sub

' Sub to set language to French
Sub Français()
    If Config.SetVar(ARES.ARES_LANGUAGE.key, "Français") Then
        ShowStatus ARES.ARES_LANGUAGE.key & " défini à Français, veuillez redémarrer."
    Else
        ShowStatus "Impossible de définir " & ARES.ARES_LANGUAGE.key & ", veuillez essayer manuellement."
    End If
End Sub
