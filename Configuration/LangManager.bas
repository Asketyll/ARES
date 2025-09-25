' Module: LangManager
' Description: This module manages translations for different languages in GUI.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: Config, ARESConfigClass, ARESConstants, ErrorHandlerClass
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
    mTranslations.Add "EN_VarResetAllSuccess", "All is reset to default value."
    mTranslations.Add "EN_VarResetError", "Unable to reset the variable."
    mTranslations.Add "EN_VarResetAllFailed", "Unable to reset all variables."
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
    mTranslations.Add "EN_AutoLengthsGUISelectElementsCaption", "Select:"
    mTranslations.Add "EN_AutoLengthsGUIOptionsCaption", "Edit auto lengths options:"
    mTranslations.Add "EN_AutoLengthsGUIOptionsMain_LabelCaption", "Enable auto Lenght"
    mTranslations.Add "EN_AutoLengthsGUIOptionsColor_LabelCaption", "Enable color update"
    mTranslations.Add "EN_AutoLengthsGUIOptionsCell_LabelCaption", "Enable ATLAS cell update"
    mTranslations.Add "EN_AutoLengthsGUIOptionsEdit_Trigger_CommandCaption", "Edit value {0}"
    mTranslations.Add "EN_AutoLengthsGUIOptionsEdit_Triggers_List_CommandCaption", "Edit triggers list"
    mTranslations.Add "EN_AutoLengthsGUIOptionsRound_LabelCaption", "Number after the decimal point:"
    mTranslations.Add "EN_AutoLengthsGUIOptionsEdit_Cells_List_CommandCaption", "Edit ATLAS cell list"
    mTranslations.Add "EN_AutoLengthsGUIOptionsEdit_Triggers_List_Error", "All your triggers should include: "
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
    mTranslations.Add "EN_ConfigExportTitle", "Export ARES Configuration"
    mTranslations.Add "EN_ConfigImportTitle", "Import ARES Configuration"
    mTranslations.Add "EN_ConfigBackupTitle", "Backup ARES Configuration"
    mTranslations.Add "EN_ConfigExportSuccess", "Configuration exported successfully to: {0}"
    mTranslations.Add "EN_ConfigImportSuccess", "Configuration imported successfully from: {0}"
    mTranslations.Add "EN_ConfigBackupSuccess", "Configuration backed up to: {0}"
    mTranslations.Add "EN_ConfigExportFailed", "Failed to export configuration"
    mTranslations.Add "EN_ConfigImportFailed", "Failed to import configuration"
    mTranslations.Add "EN_ConfigFileNotFound", "Configuration file not found: {0}"
    mTranslations.Add "EN_ConfigOverwritePrompt", "Overwrite existing modified settings?"
    mTranslations.Add "EN_ConfigImportOptions", "Import Options"
    mTranslations.Add "EN_ConfigFileFilter", "ARES Configuration Files (*.cfg)|*.cfg|All Files (*.*)|*.*"
    mTranslations.Add "EN_ConfigSelectExportLocation", "Select location to export configuration"
    mTranslations.Add "EN_ConfigSelectImportFile", "Select configuration file to import"
    mTranslations.Add "EN_ConfigOperationCancelled", "Operation cancelled by user"
    mTranslations.Add "EN_ConfigSummaryTitle", "ARES Configuration Summary"
    mTranslations.Add "EN_ConfigImportedCount", "Import completed: {0} imported, {1} skipped"
    
    ' Add French translations
    mTranslations.Add "FR_VarResetSuccess", "Réinitialisé à la valeur par défaut: {0}"
    mTranslations.Add "FR_VarResetAllSuccess", "Toutes les variables ont été remises à leur valeur par défaut."
    mTranslations.Add "FR_VarResetError", "Impossible de réinitialiser la variable."
    mTranslations.Add "FR_VarResetAllFailed", "Impossible de réinitialiser les variables."
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
    mTranslations.Add "FR_AutoLengthsGUISelectElementsCaption", "Sélectionner:"
    mTranslations.Add "FR_AutoLengthsGUIOptionsCaption", "Modifier les options de longueurs automatiques :"
    mTranslations.Add "FR_AutoLengthsGUIOptionsMain_LabelCaption", "Activer les longueurs auto."
    mTranslations.Add "FR_AutoLengthsGUIOptionsColor_LabelCaption", "Activer la MAJ de la couleur"
    mTranslations.Add "FR_AutoLengthsGUIOptionsCell_LabelCaption", "Activer la MAJ des cellules ATLAS"
    mTranslations.Add "FR_AutoLengthsGUIOptionsEdit_Trigger_CommandCaption", "Editer la valeur {0}"
    mTranslations.Add "FR_AutoLengthsGUIOptionsEdit_Triggers_List_CommandCaption", "Editer la liste des déclencheurs"
    mTranslations.Add "FR_AutoLengthsGUIOptionsRound_LabelCaption", "Nombre après la virgule:"
    mTranslations.Add "FR_AutoLengthsGUIOptionsEdit_Cells_List_CommandCaption", "Editer la liste des cellules ATLAS"
    mTranslations.Add "FR_AutoLengthsGUIOptionsEdit_Triggers_List_Error", "Tout vos déclencheurs doivent comporter: "
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
    mTranslations.Add "FR_ConfigExportTitle", "Exporter la Configuration ARES"
    mTranslations.Add "FR_ConfigImportTitle", "Importer la Configuration ARES"
    mTranslations.Add "FR_ConfigBackupTitle", "Sauvegarder la Configuration ARES"
    mTranslations.Add "FR_ConfigExportSuccess", "Configuration exportée avec succès vers: {0}"
    mTranslations.Add "FR_ConfigImportSuccess", "Configuration importée avec succès depuis: {0}"
    mTranslations.Add "FR_ConfigBackupSuccess", "Configuration sauvegardée vers: {0}"
    mTranslations.Add "FR_ConfigExportFailed", "Échec de l'export de la configuration"
    mTranslations.Add "FR_ConfigImportFailed", "Échec de l'import de la configuration"
    mTranslations.Add "FR_ConfigFileNotFound", "Fichier de configuration introuvable: {0}"
    mTranslations.Add "FR_ConfigOverwritePrompt", "Écraser les paramètres modifiés existants?"
    mTranslations.Add "FR_ConfigImportOptions", "Options d'Import"
    mTranslations.Add "FR_ConfigFileFilter", "Fichiers de Configuration ARES (*.cfg)|*.cfg|Tous les Fichiers (*.*)|*.*"
    mTranslations.Add "FR_ConfigSelectExportLocation", "Sélectionnez l'emplacement pour exporter la configuration"
    mTranslations.Add "FR_ConfigSelectImportFile", "Sélectionnez le fichier de configuration à importer"
    mTranslations.Add "FR_ConfigOperationCancelled", "Opération annulée par l'utilisateur"
    mTranslations.Add "FR_ConfigSummaryTitle", "Résumé de la Configuration ARES"
    mTranslations.Add "FR_ConfigImportedCount", "Import terminé: {0} importées, {1} ignorées"
    
    IsInit = True
    Exit Sub

ErrorHandler:
    IsInit = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "LangManager.InitializeTranslations"
    MsgBox "An error occurred while initializing translations.", vbOKOnly
End Sub

' Get translation for specified key with optional parameter substitution
' Returns localized string based on user language preference
Public Function GetTranslation(StrKey As String, ParamArray params() As Variant) As String
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If Not IsInit Then
        GetTranslation = "[Translation system not initialized] " & StrKey
        Exit Function
    End If
    
    If Len(Trim(StrKey)) = 0 Then
        GetTranslation = "[Empty translation key]"
        Exit Function
    End If
    
    Dim strBaseKey As String
    Dim strTranslatedText As String
    Dim i As Long
    
    ' Construct language-specific key
    strBaseKey = UCase(Left(mUserLanguage, 2)) & "_" & StrKey
    
    ' Try to find translation in user's language
    If mTranslations.Exists(strBaseKey) Then
        strTranslatedText = mTranslations(strBaseKey)
    Else
        ' Fallback to English if user language not available
        strBaseKey = "EN_" & StrKey
        If mTranslations.Exists(strBaseKey) Then
            strTranslatedText = mTranslations(strBaseKey)
        Else
            ' Last resort: return error message with key
            GetTranslation = "[Missing translation: " & StrKey & "]"
            Exit Function
        End If
    End If

    ' Apply parameter substitution if parameters provided
    If UBound(params) >= LBound(params) Then
        For i = LBound(params) To UBound(params)
            strTranslatedText = Replace(strTranslatedText, "{" & i & "}", CStr(params(i)))
        Next i
    End If
    
    GetTranslation = strTranslatedText
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "LangManager.GetTranslation"
    GetTranslation = "[Translation error for: " & StrKey & "]"
End Function

' Determine user's preferred language from various sources
' Priority: MicroStation config > ARES config > user prompt > default (English)
Private Function GetUserLanguage() As String
    On Error GoTo ErrorHandler
    
    Dim strLanguage As String
    
    ' First try: MicroStation CONNECT user language setting
    strLanguage = Config.GetVar("CONNECTUSER_LANGUAGE")
    If strLanguage <> "" And strLanguage <> ARESConstants.ARES_NAVD Then
        GetUserLanguage = strLanguage
        Exit Function
    End If
    
    ' Second try: ARES configuration
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    If ARESConfig.ARES_LANGUAGE.Value <> "" Then
        GetUserLanguage = ARESConfig.ARES_LANGUAGE.Value
        Exit Function
    End If
    
    ' Third try: Prompt user for language selection
    strLanguage = PromptForLanguageSelection()
    If strLanguage <> "" Then
        GetUserLanguage = strLanguage
        Exit Function
    End If
    
    ' Default fallback
    GetUserLanguage = "English"
    Exit Function

ErrorHandler:
    GetUserLanguage = "English"
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "LangManager.GetUserLanguage"
End Function

' Prompt user to select their preferred language
Private Function PromptForLanguageSelection() As String
    On Error GoTo ErrorHandler
    
    Dim strPrompt As String
    Dim varLang As Variant
    
    strPrompt = "Language Detection Failed" & vbCrLf & vbCrLf & _
                "Unable to detect your preferred language." & vbCrLf & _
                "Please set the ARES_Language environment variable." & vbCrLf & vbCrLf & _
                "Supported languages:" & vbCrLf
    
    ' Add supported languages to prompt
    For Each varLang In mSupportedLanguages
        strPrompt = strPrompt & "• " & varLang & vbCrLf
    Next varLang
    
    strPrompt = strPrompt & vbCrLf & "Available commands:" & vbCrLf & _
                "• macro vba run [ARES]English" & vbCrLf & _
                "• macro vba run [ARES]Français"
    
    MsgBox strPrompt, vbInformation + vbOKOnly, "ARES Language Configuration"
    
    PromptForLanguageSelection = "" ' User must set manually
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "LangManager.PromptForLanguageSelection"
    PromptForLanguageSelection = ""
End Function