' Module: LangManager
' Description: This module manages translations for different languages in GUI.
' License: This project is licensed under the AGPL-3.0.
' Dependencies: Config, ARESConfigClass, ARESConstants, ErrorHandlerClass
Option Explicit

Private moSupportedLanguages As Collection
Private moTranslations As Object
Private msUserLanguage As String
Public IsInit As Boolean

' Initialize translations and supported languages
Sub InitializeTranslations()
    On Error GoTo ErrorHandler
    Set moSupportedLanguages = New Collection
    Set moTranslations = CreateObject("Scripting.Dictionary")
    
    ' Add supported languages to the collection
    moSupportedLanguages.Add "English"
    moSupportedLanguages.Add "Français"

    ' Initialize user language
    msUserLanguage = GetUserLanguage()
    
    ' Add English translations
    moTranslations.Add "EN_VarResetSuccess", "Reset to default value: {0}"
    moTranslations.Add "EN_VarResetAllSuccess", "All is reset to default value."
    moTranslations.Add "EN_VarResetError", "Unable to reset the variable."
    moTranslations.Add "EN_VarResetAllFailed", "Unable to reset all variables."
    moTranslations.Add "EN_VarRemoveConfirm", "Do you really want to remove the variable {0} ?"
    moTranslations.Add "EN_VarRemoveSuccess", "Removed."
    moTranslations.Add "EN_VarRemoveError", "Unable to remove the variable."
    moTranslations.Add "EN_VarKeyNotFound", "Key not found in the collection: {0}"
    moTranslations.Add "EN_VarInvalidArgument", "Invalid argument type."
    moTranslations.Add "EN_VarInitializeMSVarfailed", "ARES Config with MS Vars failed."
    moTranslations.Add "EN_VarKeyNotInCollection", "The variable: {0} is not known."
    moTranslations.Add "EN_VarsRemoveConfirm", "Do you really want to remove all variables ? This action is irreversible."
    moTranslations.Add "EN_BootUserLangInit", "User language initialized."
    moTranslations.Add "EN_BootMSVarsInit", "Variable management initialized."
    moTranslations.Add "EN_BootMSVarsMissing", "Variable management is missing."
    moTranslations.Add "EN_BootFail", "Error in automatic loading of VBA."
    moTranslations.Add "EN_LangFail", "Translation not found for key: "
    moTranslations.Add "EN_LengthRoundError", "Rounding value unauthorized: {0}"
    moTranslations.Add "EN_LengthElementTypeNotSupportedByInterface", "The element: {0} is an element of type: {1}, it is not supported by the GetElementLength interface."
    moTranslations.Add "EN_DGNOpenCloseEventsInitialized", "Track events element initialized."
    moTranslations.Add "EN_DGNOpenCloseInitError", "Error initializing DGN Open/Close events: "
    moTranslations.Add "EN_AutoLengthsGUIInvalidSelectedElement", "The selected item is invalid."
    moTranslations.Add "EN_AutoLengthsGUISelectElementsCaption", "Select:"
    moTranslations.Add "EN_AutoLengthsGUIOptionsCaption", "Edit auto lengths options:"
    moTranslations.Add "EN_AutoLengthsGUIOptionsMain_LabelCaption", "Enable auto length"
    moTranslations.Add "EN_AutoLengthsGUIOptionsColor_LabelCaption", "Update color."
    moTranslations.Add "EN_AutoLengthsGUIOptionsOnly_Color_LabelCaption", "Update color without length."
    moTranslations.Add "EN_AutoLengthsGUIOptionsCell_LabelCaption", "Enable ATLAS cell update"
    moTranslations.Add "EN_AutoLengthsGUIOptionsEdit_Trigger_CommandCaption", "Edit value {0}"
    moTranslations.Add "EN_AutoLengthsGUIOptionsEdit_Triggers_List_CommandCaption", "Edit triggers list"
    moTranslations.Add "EN_AutoLengthsGUIOptionsRound_LabelCaption", "Number after the decimal point:"
    moTranslations.Add "EN_AutoLengthsGUIOptionsEdit_Cells_List_CommandCaption", "Edit ATLAS cell list"
    moTranslations.Add "EN_AutoLengthsGUIOptionsEdit_Triggers_List_Error", "All your triggers should include: {0}"
    moTranslations.Add "EN_AutoLengthsInitError", "Error initializing AutoLengths: "
    moTranslations.Add "EN_AutoLengthsCalculationError", "Error calculating lengths: "
    moTranslations.Add "EN_AutoLengthsUpdateError", "An error occurred while updating lengths: "
    moTranslations.Add "EN_AutoLengthsShowFormError", "Error showing element selection form: "
    moTranslations.Add "EN_AutoLengthsSelectionError", "Error selecting element: "
    moTranslations.Add "EN_AutoLengthsSetTriggerError", "Error setting trigger: "
    moTranslations.Add "EN_AutoLengthsAddTriggerError", "Error adding trigger: "
    moTranslations.Add "EN_AutoLengthsResetTriggerError", "Error resetting trigger: "
    moTranslations.Add "EN_AutoLengthsNoValidElement", "No valid element selected"
    moTranslations.Add "EN_AutoLengthsSelectanelementC", "Select an element"
    moTranslations.Add "EN_AutoLengthsSelectanelementP", "Select a valid and linked element"
    moTranslations.Add "EN_RegionSplitSelectRegionC", "Split region"
    moTranslations.Add "EN_RegionSplitSelectRegionP", "Click on the edge of a closed region to split it"
    moTranslations.Add "EN_RegionSplitNoRegion", "No valid region selected"
    moTranslations.Add "EN_RegionSplitClickNotOnEdge", "Click is not on the region edge"
    moTranslations.Add "EN_RegionSplitCannotSplit", "Cannot split this region here"
    moTranslations.Add "EN_ConfigExportTitle", "Export ARES Configuration"
    moTranslations.Add "EN_ConfigImportTitle", "Import ARES Configuration"
    moTranslations.Add "EN_ConfigBackupTitle", "Backup ARES Configuration"
    moTranslations.Add "EN_ConfigExportSuccess", "Configuration exported successfully to: {0}"
    moTranslations.Add "EN_ConfigImportSuccess", "Configuration imported successfully from: {0}"
    moTranslations.Add "EN_ConfigBackupSuccess", "Configuration backed up to: {0}"
    moTranslations.Add "EN_ConfigExportFailed", "Failed to export configuration"
    moTranslations.Add "EN_ConfigImportFailed", "Failed to import configuration"
    moTranslations.Add "EN_ConfigFileNotFound", "Configuration file not found: {0}"
    moTranslations.Add "EN_ConfigOverwritePrompt", "Overwrite existing modified settings?"
    moTranslations.Add "EN_ConfigImportOptions", "Import Options"
    moTranslations.Add "EN_ConfigFileFilter", "ARES Configuration Files (*.cfg)|*.cfg|All Files (*.*)|*.*"
    moTranslations.Add "EN_ConfigSelectExportLocation", "Select location to export configuration"
    moTranslations.Add "EN_ConfigSelectImportFile", "Select configuration file to import"
    moTranslations.Add "EN_ConfigOperationCancelled", "Operation cancelled by user"
    moTranslations.Add "EN_ConfigSummaryTitle", "ARES Configuration Summary"
    moTranslations.Add "EN_ConfigImportedCount", "Import completed: {0} imported, {1} skipped"
    moTranslations.Add "EN_ZoningGUIOptionsCaption", "Edit zoning options:"
    moTranslations.Add "EN_ZoningGUIOptionsEditLevels_CommandCaption", "Edit source levels"
    moTranslations.Add "EN_ZoningGUIOptionsDistance_LabelCaption", "Distance:"
    moTranslations.Add "EN_ZoningGUIOptionsEditOutputLevel_CommandCaption", "Edit output level ({0})"
    moTranslations.Add "EN_ZoningGUIOptionsOutputStyle_LabelCaption", "Output style:"
    moTranslations.Add "EN_ZoningGUIOptionsEditColor_CommandCaption", "Edit Color"
    moTranslations.Add "EN_ZoningGUIOptionsWeight_LabelCaption", "Weight:"
    moTranslations.Add "EN_ZoningGUIOptionsDistanceError", "Distance must be a positive number."
    moTranslations.Add "EN_OutlineGUIOptionsCaption", "Edit outline options:"
    moTranslations.Add "EN_OutlineGUIOptionsEditLevels_CommandCaption", "Edit source levels"
    moTranslations.Add "EN_OutlineGUIOptionsDistance_LabelCaption", "Distance:"
    moTranslations.Add "EN_OutlineGUIOptionsEditOutputLevel_CommandCaption", "Edit output level ({0})"
    moTranslations.Add "EN_OutlineGUIOptionsOutputStyle_LabelCaption", "Output style:"
    moTranslations.Add "EN_OutlineGUIOptionsEditColor_CommandCaption", "Edit Color"
    moTranslations.Add "EN_OutlineGUIOptionsWeight_LabelCaption", "Weight:"
    moTranslations.Add "EN_OutlineGUIOptionsDistanceError", "Distance must be a positive number."
    moTranslations.Add "EN_OutlineDistanceInvalid", "ARES: ARES_Outline_Distance invalid or empty — RunOutline aborted"
    moTranslations.Add "EN_OutlineLevelEmpty", "ARES: ARES_Outline_Level empty — RunOutline aborted"
    ' --- Messaging retrofit: generic command failure (detail goes to the .log) ---
    moTranslations.Add "EN_CommandFailed", "{0} failed"
    ' --- Language switch ---
    moTranslations.Add "EN_LanguageChanged", "ARES language set — please restart MicroStation."
    moTranslations.Add "EN_LanguageChangeFailed", "Unable to set ARES language — set ARES_Language manually."
    ' --- Change tracking (bulk suspend/resume) ---
    moTranslations.Add "EN_ChangeTrackingAlreadySuspended", "ARES: Change tracking already suspended"
    moTranslations.Add "EN_ChangeTrackingSuspended", "ARES: Change tracking suspended — perform the bulk operation, then resume"
    moTranslations.Add "EN_ChangeTrackingNoHandler", "ARES: No change handler to suspend"
    ' --- Zone export (user-facing results; progress steps go to the .log) ---
    moTranslations.Add "EN_ZoneExportNoActiveModel", "ARES: Zone export — no active model reference"
    moTranslations.Add "EN_ZoneExportLevelNotConfigured", "ARES: Zone export — zone level not configured"
    moTranslations.Add "EN_ZoneExportLevelNotFound", "ARES: Zone export — zone level not found: {0}"
    moTranslations.Add "EN_ZoneExportCancelled", "ARES: Zone export — cancelled"
    moTranslations.Add "EN_ZoneExportNoZones", "ARES: Zone export — no zones on level {0}"
    moTranslations.Add "EN_ZoneExportComplete", "ARES: Zone export complete — {0} elements, {1} groups ({2})"
    moTranslations.Add "EN_ZoneExportCompletePerZone", "ARES: Zone export complete — {0} elements, {1} rows per zone ({2})"
    moTranslations.Add "EN_ZoneExportFailed", "ARES: Zone export failed"
    moTranslations.Add "EN_ZoneExportFilterLevelsIgnored", "ARES: Zone export — filter level(s) ignored (not found): {0}"
    moTranslations.Add "EN_ZoneExportZonePropertyInvalid", "ARES: Zone export — zone property invalid, using zone index"
    ' --- Property Tagging (custom-property) options GUI ---
    moTranslations.Add "EN_PropertyTaggingGUIOptionsCaption", "Edit custom-property options:"
    moTranslations.Add "EN_PropertyTaggingGUIOptionsMain_LabelCaption", "Auto-attach on create / modify"
    moTranslations.Add "EN_PropertyTaggingGUIOptionsEditList_CommandCaption", "Edit property list"
    moTranslations.Add "EN_PropertyTaggingGUIOptionsEditRules_CommandCaption", "Edit rules"
    moTranslations.Add "EN_ZoningNoBufferCreated", "No buffer could be created for any of the {0} element(s) found."
    moTranslations.Add "EN_ZoningSomeBuffersFailed", "{0} of {1} element(s) could not be buffered and were skipped."
    moTranslations.Add "EN_ZoneExportGUIOptionsCaption", "Edit zone export options:"
    moTranslations.Add "EN_ZoneExportGUIOptionsEdit_Level_Region_CommandCaption", "Edit zone level"
    moTranslations.Add "EN_ZoneExportGUIOptionsEdit_Level_Candidate_CommandCaption", "Edit filter level"
    moTranslations.Add "EN_ZoneExportGUIOptionsRound_LabelCaption", "Decimal places:"
    moTranslations.Add "EN_ZoneExportGUIOptionsUse_Dialog_LabelCaption", "Prompt save location"
    moTranslations.Add "EN_WikiOpenFailed", "Failed to open ARES wiki"
    moTranslations.Add "EN_UpdateAvailableTitle", "ARES - Update Available"
    moTranslations.Add "EN_UpdateAvailableQuestion", "A new version of ARES is available. Do you want to update?"
    moTranslations.Add "EN_UpdateBtnYes", "Yes"
    moTranslations.Add "EN_UpdateBtnNo", "No"
    moTranslations.Add "EN_UpdateBtnIgnoreAll", "Ignore all"
    moTranslations.Add "EN_UpdateDownloading", "Downloading update..."
    moTranslations.Add "EN_UpdateDownloadFailed", "Failed to download the update. Please visit the GitHub releases page."
    moTranslations.Add "EN_UpdateCheckFailed", "ARES: Update check failed. Check your network connection."
    moTranslations.Add "EN_UpdateAlreadyUpToDate", "ARES is up to date."
    moTranslations.Add "EN_ChangeTrackingResumed", "ARES: Change tracking resumed after bulk operation"
    moTranslations.Add "EN_ChangeTrackingResumeWarning", "ARES: WARNING - change tracking NOT attached after bulk resume"
    ' --- Story 8-1: shared form-UX baseline (FormUXHelper) ---
    moTranslations.Add "EN_FormFinishEditFirst", "Finish the current edit, or press Esc to cancel."
    moTranslations.Add "EN_FormResetDefaultsCaption", "Restore defaults"
    moTranslations.Add "EN_FormDefaultsRestored", "Default options restored."
    moTranslations.Add "EN_FormPositionsReset", "Window positions reset."
    moTranslations.Add "EN_UpdateBtnSkipVersion", "Skip this version"
    moTranslations.Add "EN_UpdateBtnYesTip", "Download and install the new version now."
    moTranslations.Add "EN_UpdateBtnSkipVersionTip", "Do not remind me about this version again (newer versions will still be announced)."
    moTranslations.Add "EN_UpdateBtnIgnoreAllTip", "Mute ALL future update prompts."
    moTranslations.Add "EN_ZoneExportGUIOptionsGroupBy_LabelCaption", "Group by"
    moTranslations.Add "EN_ZoneExportGroupByStyle", "Style"
    moTranslations.Add "EN_ZoneExportGroupByLevel", "Level"
    moTranslations.Add "EN_ZoneExportGroupByColor", "Color"
    moTranslations.Add "EN_ZoneExportGUIOptionsPerZone_LabelCaption", "Break down by property"
    moTranslations.Add "EN_ZoneExportGUIOptionsZoneProperty_LabelCaption", "Property used:"
    ' --- Story 8-2: restore-defaults tooltip + element-picker OK/Cancel ---
    moTranslations.Add "EN_FormResetDefaultsTip", "Reset every option on this panel to its default value."
    moTranslations.Add "EN_AutoLengthsGUISelectElementsOK_CommandCaption", "Select"
    moTranslations.Add "EN_AutoLengthsGUISelectElementsCancel_CommandCaption", "Cancel"
    ' Tooltips (ControlTipText) - Property Tagging
    moTranslations.Add "EN_PropertyTaggingGUIOptionsMain_LabelTip", "Attach ARES custom properties automatically when elements are created or modified."
    moTranslations.Add "EN_PropertyTaggingGUIOptionsEditList_CommandTip", "Pipe-separated (|) list of custom-property names. Example: Commune|Coupe_Type"
    moTranslations.Add "EN_PropertyTaggingGUIOptionsEditRules_CommandTip", "Rules: level[:type]=prop|prop ; ...  Example: WALLS=Commune|Coupe_Type ; DOORS:Cell=Commune"
    ' Tooltips - Zoning
    moTranslations.Add "EN_ZoningGUIOptionsEditLevels_CommandTip", "Source levels to process (pipe-separated |)."
    moTranslations.Add "EN_ZoningGUIOptionsDistance_LabelTip", "Buffer distance in master units. Must be a positive number."
    moTranslations.Add "EN_ZoningGUIOptionsEditOutputLevel_CommandTip", "Level the output zones are created on."
    moTranslations.Add "EN_ZoningGUIOptionsEditColor_CommandTip", "Pick the output color (MicroStation color index)."
    moTranslations.Add "EN_ZoningGUIOptionsColor_SwatchTip", "Current output color."
    moTranslations.Add "EN_ZoningGUIOptionsOutputStyle_LabelTip", "Output line style (index or named style)."
    moTranslations.Add "EN_ZoningGUIOptionsWeight_LabelTip", "Output line weight (0-31)."
    ' Tooltips - Outline
    moTranslations.Add "EN_OutlineGUIOptionsEditLevels_CommandTip", "Source levels to process (pipe-separated |)."
    moTranslations.Add "EN_OutlineGUIOptionsDistance_LabelTip", "Buffer distance in master units. Must be a positive number."
    moTranslations.Add "EN_OutlineGUIOptionsEditOutputLevel_CommandTip", "Level the output zones are created on."
    moTranslations.Add "EN_OutlineGUIOptionsEditColor_CommandTip", "Pick the output color (MicroStation color index)."
    moTranslations.Add "EN_OutlineGUIOptionsColor_SwatchTip", "Current output color."
    moTranslations.Add "EN_OutlineGUIOptionsOutputStyle_LabelTip", "Output line style (index or named style)."
    moTranslations.Add "EN_OutlineGUIOptionsWeight_LabelTip", "Output line weight (0-31)."
    ' Tooltips - Zone Export
    moTranslations.Add "EN_ZoneExportGUIOptionsEdit_Level_Region_CommandTip", "Level holding the zone regions to measure inside."
    moTranslations.Add "EN_ZoneExportGUIOptionsEdit_Level_Candidate_CommandTip", "Restrict measured elements to these level(s), pipe-separated (|). Empty = all levels."
    moTranslations.Add "EN_ZoneExportGUIOptionsGroupBy_LabelTip", "How exported rows are grouped: line style, level, color, or a per-zone breakdown by custom property."
    moTranslations.Add "EN_ZoneExportGUIOptionsPerZone_LabelTip", "Split each group by zone; label each zone with the chosen zone property."
    moTranslations.Add "EN_ZoneExportGUIOptionsZoneProperty_LabelTip", "Custom property whose value labels each zone in the Zone column (when 'Break down by property' is on). Empty or invalid = zones numbered Zone 1, Zone 2, ..."
    moTranslations.Add "EN_ZoneExportGUIOptionsRound_LabelTip", "Decimal places for exported lengths (0-10)."
    moTranslations.Add "EN_ZoneExportGUIOptionsUse_Dialog_LabelTip", "When on, prompt for the save location; otherwise auto-name the file."
    ' Tooltips - Auto Lengths
    moTranslations.Add "EN_AutoLengthsGUIOptionsMain_LabelTip", "Automatically write the linked-geometry length into text triggers."
    moTranslations.Add "EN_AutoLengthsGUIOptionsColor_LabelTip", "Also sync the text color from the linked geometry."
    moTranslations.Add "EN_AutoLengthsGUIOptionsOnly_Color_LabelTip", "Update only the color, not the length value."
    moTranslations.Add "EN_AutoLengthsGUIOptionsCell_LabelTip", "Rebuild ATLAS label cells after a text edit."
    moTranslations.Add "EN_AutoLengthsGUIOptionsEdit_Trigger_CommandTip", "The token replaced by the length (must appear in every trigger)."
    moTranslations.Add "EN_AutoLengthsGUIOptionsEdit_Triggers_List_CommandTip", "Trigger patterns (pipe-separated |). Each must contain the trigger token."
    moTranslations.Add "EN_AutoLengthsGUIOptionsEdit_Cells_List_CommandTip", "ATLAS label cell names (pipe-separated |)."
    moTranslations.Add "EN_AutoLengthsGUIOptionsRound_LabelTip", "Decimals kept in the written length."
    moTranslations.Add "EN_AutoLengthsGUISelectElementsListTip", "Double-click or press Enter to pick; Esc to cancel."

    ' Add French translations
    moTranslations.Add "FR_VarResetSuccess", "Réinitialisé à la valeur par défaut: {0}"
    moTranslations.Add "FR_VarResetAllSuccess", "Toutes les variables ont été remises à leur valeur par défaut."
    moTranslations.Add "FR_VarResetError", "Impossible de réinitialiser la variable."
    moTranslations.Add "FR_VarResetAllFailed", "Impossible de réinitialiser les variables."
    moTranslations.Add "FR_VarRemoveConfirm", "Voulez-vous vraiment supprimer la variable {0}?"
    moTranslations.Add "FR_VarRemoveSuccess", "Supprimé."
    moTranslations.Add "FR_VarRemoveError", "Impossible de supprimer la variable."
    moTranslations.Add "FR_VarKeyNotFound", "Clé introuvable dans la collection: {0}"
    moTranslations.Add "FR_VarInvalidArgument", "Type d'argument non valide."
    moTranslations.Add "FR_VarInitializeMSVarfailed", "ARES Config avec MS Vars à échoué."
    moTranslations.Add "FR_VarKeyNotInCollection", "La variable: {0} n'est pas reconnue."
    moTranslations.Add "FR_VarsRemoveConfirm", "Voulez-vous vraiment supprimer toutes les variables ? Cette action est irréversible."
    moTranslations.Add "FR_BootUserLangInit", "Langage utilisateur initialisé."
    moTranslations.Add "FR_BootMSVarsInit", "Gestion des variables initialisées."
    moTranslations.Add "FR_BootMSVarsMissing", "Gestion des variables manquante."
    moTranslations.Add "FR_BootFail", "Erreur lors du chargement automatique de VBA."
    moTranslations.Add "FR_LangFail", "Traduction introuvable pour la clé: "
    moTranslations.Add "FR_LengthRoundError", "Valeur d'arrondi interdite : {0}"
    moTranslations.Add "FR_LengthElementTypeNotSupportedByInterface", "L'élément: {0} est un élément de type: {1}, il n'est pas géré par l'interface GetElementLength."
    moTranslations.Add "FR_DGNOpenCloseEventsInitialized", "Evénements de suivi d'objet initialisé."
    moTranslations.Add "FR_DGNOpenCloseInitError", "Erreur lors de l'initialisation des événements d'ouverture/fermeture DGN: "
    moTranslations.Add "FR_AutoLengthsGUIInvalidSelectedElement", "L'élément sélectionné n'est pas valide."
    moTranslations.Add "FR_AutoLengthsGUISelectElementsCaption", "Sélectionner:"
    moTranslations.Add "FR_AutoLengthsGUIOptionsCaption", "Modifier les options de longueurs automatiques :"
    moTranslations.Add "FR_AutoLengthsGUIOptionsMain_LabelCaption", "Activer les longueurs auto."
    moTranslations.Add "FR_AutoLengthsGUIOptionsColor_LabelCaption", "MAJ de la couleur."
    moTranslations.Add "FR_AutoLengthsGUIOptionsOnly_Color_LabelCaption", "MAJ de la couleur sans longueur."
    moTranslations.Add "FR_AutoLengthsGUIOptionsCell_LabelCaption", "Activer la MAJ des cellules ATLAS"
    moTranslations.Add "FR_AutoLengthsGUIOptionsEdit_Trigger_CommandCaption", "Editer la valeur {0}"
    moTranslations.Add "FR_AutoLengthsGUIOptionsEdit_Triggers_List_CommandCaption", "Editer la liste des déclencheurs"
    moTranslations.Add "FR_AutoLengthsGUIOptionsRound_LabelCaption", "Nombre après la virgule:"
    moTranslations.Add "FR_AutoLengthsGUIOptionsEdit_Cells_List_CommandCaption", "Editer la liste des cellules ATLAS"
    moTranslations.Add "FR_AutoLengthsGUIOptionsEdit_Triggers_List_Error", "Tous vos déclencheurs doivent comporter: {0}"
    moTranslations.Add "FR_AutoLengthsInitError", "Erreur lors de l'initialisation d'AutoLengths: "
    moTranslations.Add "FR_AutoLengthsCalculationError", "Erreur lors du calcul des longueurs: "
    moTranslations.Add "FR_AutoLengthsUpdateError", "Une erreur s'est produite lors de la mise à jour des longueurs: "
    moTranslations.Add "FR_AutoLengthsShowFormError", "Erreur lors de l'affichage du formulaire de sélection d'élément: "
    moTranslations.Add "FR_AutoLengthsSelectionError", "Erreur lors de la sélection de l'élément: "
    moTranslations.Add "FR_AutoLengthsSetTriggerError", "Erreur lors de la définition du déclencheur: "
    moTranslations.Add "FR_AutoLengthsAddTriggerError", "Erreur lors de l'ajout du déclencheur: "
    moTranslations.Add "FR_AutoLengthsResetTriggerError", "Erreur lors de la réinitialisation du déclencheur: "
    moTranslations.Add "FR_AutoLengthsNoValidElement", "Aucun élément valide sélectionné"
    moTranslations.Add "FR_AutoLengthsSelectanelementC", "Sélectionner un élément"
    moTranslations.Add "FR_AutoLengthsSelectanelementP", "Sélectionner un élément valide et groupé"
    moTranslations.Add "FR_RegionSplitSelectRegionC", "Diviser une région"
    moTranslations.Add "FR_RegionSplitSelectRegionP", "Cliquer sur le bord d'une région fermée pour la diviser"
    moTranslations.Add "FR_RegionSplitNoRegion", "Aucune région valide sélectionnée"
    moTranslations.Add "FR_RegionSplitClickNotOnEdge", "Le clic n'est pas sur le bord de la région"
    moTranslations.Add "FR_RegionSplitCannotSplit", "Impossible de diviser cette région ici"
    moTranslations.Add "FR_ConfigExportTitle", "Exporter la Configuration ARES"
    moTranslations.Add "FR_ConfigImportTitle", "Importer la Configuration ARES"
    moTranslations.Add "FR_ConfigBackupTitle", "Sauvegarder la Configuration ARES"
    moTranslations.Add "FR_ConfigExportSuccess", "Configuration exportée avec succès vers: {0}"
    moTranslations.Add "FR_ConfigImportSuccess", "Configuration importée avec succès depuis: {0}"
    moTranslations.Add "FR_ConfigBackupSuccess", "Configuration sauvegardée vers: {0}"
    moTranslations.Add "FR_ConfigExportFailed", "Échec de l'export de la configuration"
    moTranslations.Add "FR_ConfigImportFailed", "Échec de l'import de la configuration"
    moTranslations.Add "FR_ConfigFileNotFound", "Fichier de configuration introuvable: {0}"
    moTranslations.Add "FR_ConfigOverwritePrompt", "Écraser les paramètres modifiés existants?"
    moTranslations.Add "FR_ConfigImportOptions", "Options d'Import"
    moTranslations.Add "FR_ConfigFileFilter", "Fichiers de Configuration ARES (*.cfg)|*.cfg|Tous les Fichiers (*.*)|*.*"
    moTranslations.Add "FR_ConfigSelectExportLocation", "Sélectionnez l'emplacement pour exporter la configuration"
    moTranslations.Add "FR_ConfigSelectImportFile", "Sélectionnez le fichier de configuration à importer"
    moTranslations.Add "FR_ConfigOperationCancelled", "Opération annulée par l'utilisateur"
    moTranslations.Add "FR_ConfigSummaryTitle", "Résumé de la Configuration ARES"
    moTranslations.Add "FR_ConfigImportedCount", "Import terminé: {0} importées, {1} ignorées"
    moTranslations.Add "FR_ZoningGUIOptionsCaption", "Modifier les options de zonage :"
    moTranslations.Add "FR_ZoningGUIOptionsEditLevels_CommandCaption", "Modifier les niveaux sources"
    moTranslations.Add "FR_ZoningGUIOptionsDistance_LabelCaption", "Distance :"
    moTranslations.Add "FR_ZoningGUIOptionsEditOutputLevel_CommandCaption", "Modifier le niveau de sortie ({0})"
    moTranslations.Add "FR_ZoningGUIOptionsOutputStyle_LabelCaption", "Style :"
    moTranslations.Add "FR_ZoningGUIOptionsEditColor_CommandCaption", "Modifier la couleur"
    moTranslations.Add "FR_ZoningGUIOptionsWeight_LabelCaption", "Épaisseur :"
    moTranslations.Add "FR_ZoningGUIOptionsDistanceError", "La distance doit être un nombre positif."
    moTranslations.Add "FR_OutlineGUIOptionsCaption", "Modifier les options de contour :"
    moTranslations.Add "FR_OutlineGUIOptionsEditLevels_CommandCaption", "Modifier les niveaux sources"
    moTranslations.Add "FR_OutlineGUIOptionsDistance_LabelCaption", "Distance :"
    moTranslations.Add "FR_OutlineGUIOptionsEditOutputLevel_CommandCaption", "Modifier le niveau de sortie ({0})"
    moTranslations.Add "FR_OutlineGUIOptionsOutputStyle_LabelCaption", "Style :"
    moTranslations.Add "FR_OutlineGUIOptionsEditColor_CommandCaption", "Modifier la couleur"
    moTranslations.Add "FR_OutlineGUIOptionsWeight_LabelCaption", "Épaisseur :"
    moTranslations.Add "FR_OutlineGUIOptionsDistanceError", "La distance doit être un nombre positif."
    moTranslations.Add "FR_OutlineDistanceInvalid", "ARES : ARES_Outline_Distance invalide ou vide — RunOutline annulé"
    moTranslations.Add "FR_OutlineLevelEmpty", "ARES : ARES_Outline_Level vide — RunOutline annulé"
    ' --- Messaging retrofit: generic command failure (detail goes to the .log) ---
    moTranslations.Add "FR_CommandFailed", "{0} a échoué"
    ' --- Language switch ---
    moTranslations.Add "FR_LanguageChanged", "Langue ARES définie — veuillez redémarrer MicroStation."
    moTranslations.Add "FR_LanguageChangeFailed", "Impossible de définir la langue ARES — définissez ARES_Language manuellement."
    ' --- Change tracking (bulk suspend/resume) ---
    moTranslations.Add "FR_ChangeTrackingAlreadySuspended", "ARES : Suivi des modifications déjà suspendu"
    moTranslations.Add "FR_ChangeTrackingSuspended", "ARES : Suivi des modifications suspendu — effectuez l'opération en lot, puis reprenez"
    moTranslations.Add "FR_ChangeTrackingNoHandler", "ARES : Aucun gestionnaire de suivi à suspendre"
    ' --- Zone export (user-facing results; progress steps go to the .log) ---
    moTranslations.Add "FR_ZoneExportNoActiveModel", "ARES : Export de zone — aucun modèle actif"
    moTranslations.Add "FR_ZoneExportLevelNotConfigured", "ARES : Export de zone — niveau de zone non configuré"
    moTranslations.Add "FR_ZoneExportLevelNotFound", "ARES : Export de zone — niveau de zone introuvable : {0}"
    moTranslations.Add "FR_ZoneExportCancelled", "ARES : Export de zone — annulé"
    moTranslations.Add "FR_ZoneExportNoZones", "ARES : Export de zone — aucune zone sur le niveau {0}"
    moTranslations.Add "FR_ZoneExportComplete", "ARES : Export de zone terminé — {0} éléments, {1} groupes ({2})"
    moTranslations.Add "FR_ZoneExportCompletePerZone", "ARES : Export de zone terminé — {0} éléments, {1} lignes par zone ({2})"
    moTranslations.Add "FR_ZoneExportFailed", "ARES : Échec de l'export de zone"
    moTranslations.Add "FR_ZoneExportFilterLevelsIgnored", "ARES : Export de zone — niveau(x) de filtre ignoré(s) (introuvable) : {0}"
    moTranslations.Add "FR_ZoneExportZonePropertyInvalid", "ARES : Export de zone — propriété de zone invalide, index de zone utilisé"
    ' --- Property Tagging (custom-property) options GUI ---
    moTranslations.Add "FR_PropertyTaggingGUIOptionsCaption", "Modifier les options de propriétés personnalisées :"
    moTranslations.Add "FR_PropertyTaggingGUIOptionsMain_LabelCaption", "Attache auto à la création / modification"
    moTranslations.Add "FR_PropertyTaggingGUIOptionsEditList_CommandCaption", "Modifier la liste des propriétés"
    moTranslations.Add "FR_PropertyTaggingGUIOptionsEditRules_CommandCaption", "Modifier les règles"
    moTranslations.Add "FR_ZoningNoBufferCreated", "Aucun buffer n'a pu être créé pour les {0} élément(s) trouvé(s)."
    moTranslations.Add "FR_ZoningSomeBuffersFailed", "{0} des {1} élément(s) n'ont pas pu être bufférisés et ont été ignorés."
    moTranslations.Add "FR_ZoneExportGUIOptionsCaption", "Modifier les options d'export de zone :"
    moTranslations.Add "FR_ZoneExportGUIOptionsEdit_Level_Region_CommandCaption", "Modifier le niveau de zone"
    moTranslations.Add "FR_ZoneExportGUIOptionsEdit_Level_Candidate_CommandCaption", "Modifier le niveau de filtre"
    moTranslations.Add "FR_ZoneExportGUIOptionsRound_LabelCaption", "Décimales :"
    moTranslations.Add "FR_ZoneExportGUIOptionsUse_Dialog_LabelCaption", "Demander l'emplacement d'export"
    moTranslations.Add "FR_WikiOpenFailed", "Echec de l'ouverture du wiki ARES"
    moTranslations.Add "FR_UpdateAvailableTitle", "ARES - Mise a jour disponible"
    moTranslations.Add "FR_UpdateAvailableQuestion", "Une nouvelle version d'ARES est disponible, souhaitez-vous faire la mise a jour ?"
    moTranslations.Add "FR_UpdateBtnYes", "Oui"
    moTranslations.Add "FR_UpdateBtnNo", "Non"
    moTranslations.Add "FR_UpdateBtnIgnoreAll", "Tout ignorer"
    moTranslations.Add "FR_UpdateDownloading", "Telechargement de la mise a jour..."
    moTranslations.Add "FR_UpdateDownloadFailed", "Echec du telechargement. Veuillez visiter la page des releases GitHub."
    moTranslations.Add "FR_UpdateCheckFailed", "ARES : Echec de la verification. Verifiez votre connexion reseau."
    moTranslations.Add "FR_UpdateAlreadyUpToDate", "ARES est a jour."
    moTranslations.Add "FR_ChangeTrackingResumed", "ARES : Suivi des modifications repris après l'opération en masse"
    moTranslations.Add "FR_ChangeTrackingResumeWarning", "ARES : ATTENTION - le suivi des modifications n'a PAS été réattaché après l'opération en masse"
    ' --- Story 8-1: shared form-UX baseline (FormUXHelper) ---
    moTranslations.Add "FR_FormFinishEditFirst", "Terminez la saisie en cours, ou appuyez sur Échap pour annuler."
    moTranslations.Add "FR_FormResetDefaultsCaption", "Réinitialiser"
    moTranslations.Add "FR_FormDefaultsRestored", "Options par défaut restaurées."
    moTranslations.Add "FR_FormPositionsReset", "Positions des fenêtres réinitialisées."
    moTranslations.Add "FR_UpdateBtnSkipVersion", "Ignorer cette version"
    moTranslations.Add "FR_UpdateBtnYesTip", "Télécharger et installer la nouvelle version maintenant."
    moTranslations.Add "FR_UpdateBtnSkipVersionTip", "Ne plus me rappeler cette version (les versions plus récentes seront toujours signalées)."
    moTranslations.Add "FR_UpdateBtnIgnoreAllTip", "Désactiver TOUTES les notifications de mise à jour futures."
    moTranslations.Add "FR_ZoneExportGUIOptionsGroupBy_LabelCaption", "Grouper par"
    moTranslations.Add "FR_ZoneExportGroupByStyle", "Style"
    moTranslations.Add "FR_ZoneExportGroupByLevel", "Niveau"
    moTranslations.Add "FR_ZoneExportGroupByColor", "Couleur"
    moTranslations.Add "FR_ZoneExportGUIOptionsPerZone_LabelCaption", "Répartir par propriété"
    moTranslations.Add "FR_ZoneExportGUIOptionsZoneProperty_LabelCaption", "Propriété utilisée :"
    ' --- Story 8-2 : info-bulle reinitialisation + OK/Annuler du selecteur d'elements ---
    moTranslations.Add "FR_FormResetDefaultsTip", "Réinitialise chaque option de ce panneau à sa valeur par défaut."
    moTranslations.Add "FR_AutoLengthsGUISelectElementsOK_CommandCaption", "Sélectionner"
    moTranslations.Add "FR_AutoLengthsGUISelectElementsCancel_CommandCaption", "Annuler"
    ' Tooltips (ControlTipText) - Property Tagging
    moTranslations.Add "FR_PropertyTaggingGUIOptionsMain_LabelTip", "Attache automatiquement les propriétés ARES à la création ou à la modification d'éléments."
    moTranslations.Add "FR_PropertyTaggingGUIOptionsEditList_CommandTip", "Liste de noms de propriétés séparés par | . Exemple : Commune|Coupe_Type"
    moTranslations.Add "FR_PropertyTaggingGUIOptionsEditRules_CommandTip", "Règles : niveau[:type]=prop|prop ; ...  Exemple : WALLS=Commune|Coupe_Type ; DOORS:Cell=Commune"
    ' Tooltips - Zoning
    moTranslations.Add "FR_ZoningGUIOptionsEditLevels_CommandTip", "Niveaux sources à traiter (séparés par |)."
    moTranslations.Add "FR_ZoningGUIOptionsDistance_LabelTip", "Distance de la zone tampon en unités maître. Doit être un nombre positif."
    moTranslations.Add "FR_ZoningGUIOptionsEditOutputLevel_CommandTip", "Niveau sur lequel les zones de sortie sont créées."
    moTranslations.Add "FR_ZoningGUIOptionsEditColor_CommandTip", "Choisir la couleur de sortie (index de couleur MicroStation)."
    moTranslations.Add "FR_ZoningGUIOptionsColor_SwatchTip", "Couleur de sortie actuelle."
    moTranslations.Add "FR_ZoningGUIOptionsOutputStyle_LabelTip", "Style de ligne de sortie (index ou nom de style)."
    moTranslations.Add "FR_ZoningGUIOptionsWeight_LabelTip", "Épaisseur de ligne de sortie (0-31)."
    ' Tooltips - Outline
    moTranslations.Add "FR_OutlineGUIOptionsEditLevels_CommandTip", "Niveaux sources à traiter (séparés par |)."
    moTranslations.Add "FR_OutlineGUIOptionsDistance_LabelTip", "Distance de la zone tampon en unités maître. Doit être un nombre positif."
    moTranslations.Add "FR_OutlineGUIOptionsEditOutputLevel_CommandTip", "Niveau sur lequel les zones de sortie sont créées."
    moTranslations.Add "FR_OutlineGUIOptionsEditColor_CommandTip", "Choisir la couleur de sortie (index de couleur MicroStation)."
    moTranslations.Add "FR_OutlineGUIOptionsColor_SwatchTip", "Couleur de sortie actuelle."
    moTranslations.Add "FR_OutlineGUIOptionsOutputStyle_LabelTip", "Style de ligne de sortie (index ou nom de style)."
    moTranslations.Add "FR_OutlineGUIOptionsWeight_LabelTip", "Épaisseur de ligne de sortie (0-31)."
    ' Tooltips - Zone Export
    moTranslations.Add "FR_ZoneExportGUIOptionsEdit_Level_Region_CommandTip", "Niveau contenant les régions de zone où mesurer."
    moTranslations.Add "FR_ZoneExportGUIOptionsEdit_Level_Candidate_CommandTip", "Limite les éléments mesurés à ce(s) niveau(x), séparés par |. Vide = tous les niveaux."
    moTranslations.Add "FR_ZoneExportGUIOptionsGroupBy_LabelTip", "Regroupement des lignes exportées : style, niveau, couleur, ou répartition par zone selon une propriété personnalisée."
    moTranslations.Add "FR_ZoneExportGUIOptionsPerZone_LabelTip", "Répartit chaque groupe par zone ; étiquette chaque zone avec la propriété de zone choisie."
    moTranslations.Add "FR_ZoneExportGUIOptionsZoneProperty_LabelTip", "Propriété personnalisée dont la valeur étiquette chaque zone dans la colonne Zone (quand « Répartir par propriété » est actif). Vide ou invalide = zones numérotées Zone 1, Zone 2, ..."
    moTranslations.Add "FR_ZoneExportGUIOptionsRound_LabelTip", "Décimales pour les longueurs exportées (0-10)."
    moTranslations.Add "FR_ZoneExportGUIOptionsUse_Dialog_LabelTip", "Si activé, demande l'emplacement d'export ; sinon nomme le fichier automatiquement."
    ' Tooltips - Auto Lengths
    moTranslations.Add "FR_AutoLengthsGUIOptionsMain_LabelTip", "Écrit automatiquement la longueur de la géométrie liée dans les déclencheurs texte."
    moTranslations.Add "FR_AutoLengthsGUIOptionsColor_LabelTip", "Synchronise aussi la couleur du texte depuis la géométrie liée."
    moTranslations.Add "FR_AutoLengthsGUIOptionsOnly_Color_LabelTip", "Met à jour uniquement la couleur, pas la valeur de longueur."
    moTranslations.Add "FR_AutoLengthsGUIOptionsCell_LabelTip", "Reconstruit les cellules d'étiquette ATLAS après une modification de texte."
    moTranslations.Add "FR_AutoLengthsGUIOptionsEdit_Trigger_CommandTip", "Le jeton remplacé par la longueur (doit figurer dans chaque déclencheur)."
    moTranslations.Add "FR_AutoLengthsGUIOptionsEdit_Triggers_List_CommandTip", "Motifs de déclencheurs (séparés par |). Chacun doit contenir le jeton."
    moTranslations.Add "FR_AutoLengthsGUIOptionsEdit_Cells_List_CommandTip", "Noms des cellules d'étiquette ATLAS (séparés par |)."
    moTranslations.Add "FR_AutoLengthsGUIOptionsRound_LabelTip", "Décimales conservées dans la longueur écrite."
    moTranslations.Add "FR_AutoLengthsGUISelectElementsListTip", "Double-cliquez ou appuyez sur Entrée pour choisir ; Échap pour annuler."

    IsInit = True
    Exit Sub

ErrorHandler:
    IsInit = False
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "LangManager.InitializeTranslations"
    MsgBox "An error occurred while initializing translations.", vbOKOnly
End Sub

' Get translation for specified key with optional parameter substitution
' Returns localized string based on user language preference
Public Function GetTranslation(sKey As String, ParamArray params() As Variant) As String
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If Not IsInit Then
        GetTranslation = "[Translation system not initialized] " & sKey
        Exit Function
    End If
    
    If Len(Trim(sKey)) = 0 Then
        GetTranslation = "[Empty translation key]"
        Exit Function
    End If
    
    Dim sBaseKey As String
    Dim sTranslatedText As String
    Dim i As Long
    
    ' Construct language-specific key
    sBaseKey = UCase(Left(msUserLanguage, 2)) & "_" & sKey
    
    ' Try to find translation in user's language
    If moTranslations.Exists(sBaseKey) Then
        sTranslatedText = moTranslations(sBaseKey)
    Else
        ' Fallback to English if user language not available
        sBaseKey = "EN_" & sKey
        If moTranslations.Exists(sBaseKey) Then
            sTranslatedText = moTranslations(sBaseKey)
        Else
            ' Last resort: return error message with key
            GetTranslation = "[Missing translation: " & sKey & "]"
            Exit Function
        End If
    End If

    ' Apply parameter substitution if parameters provided
    If UBound(params) >= LBound(params) Then
        For i = LBound(params) To UBound(params)
            sTranslatedText = Replace(sTranslatedText, "{" & i & "}", CStr(params(i)))
        Next i
    End If
    
    GetTranslation = sTranslatedText
    Exit Function

ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "LangManager.GetTranslation"
    GetTranslation = "[Translation error for: " & sKey & "]"
End Function

' Show a user-facing status line, translated. Self-initialises the translation system so
' callers never leak the "[not initialized]" sentinel. This is THE channel for parameter-less
' user status; for messages carrying an identifier/count call ShowStatus GetTranslation(key, args)
' directly (after ensuring init). Diagnostics/faults NEVER come here — they go to
' ErrorHandler.HandleError (the .log). See the messaging rules in project-context.md / MVBA README.
Public Sub ShowStatusT(ByVal sKey As String)
    On Error Resume Next
    If Not IsInit Then InitializeTranslations
    ShowStatus GetTranslation(sKey)
End Sub

' Return the resolved user language (e.g. "English", "Français")
' Falls back to English if the translation system has not resolved a language yet
Public Function UserLanguage() As String
    If Len(msUserLanguage) > 0 Then
        UserLanguage = msUserLanguage
    Else
        UserLanguage = "English"
    End If
End Function

' Determine user's preferred language from various sources
' Priority: MicroStation config > ARES config > user prompt > default (English)
Private Function GetUserLanguage() As String
    On Error GoTo ErrorHandler
    
    Dim sLanguage As String
    
    ' First try: MicroStation CONNECT user language setting
    sLanguage = Config.GetVar("CONNECTUSER_LANGUAGE")
    If sLanguage <> "" And sLanguage <> ARESConstants.ARES_NAVD Then
        GetUserLanguage = sLanguage
        Exit Function
    End If
    
    ' Second try: ARES configuration
    If Not ARESConfig.IsInitialized Then ARESConfig.Initialize
    If ARESConfig.ARES_LANGUAGE.Value <> "" Then
        GetUserLanguage = ARESConfig.ARES_LANGUAGE.Value
        Exit Function
    End If
    
    ' Third try: Prompt user for language selection
    sLanguage = PromptForLanguageSelection()
    If sLanguage <> "" Then
        GetUserLanguage = sLanguage
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
    
    Dim sPrompt As String
    Dim varLang As Variant
    
    sPrompt = "Language Detection Failed" & vbCrLf & vbCrLf & _
                "Unable to detect your preferred language." & vbCrLf & _
                "Please set the ARES_Language environment variable." & vbCrLf & vbCrLf & _
                "Supported languages:" & vbCrLf
    
    ' Add supported languages to prompt
    For Each varLang In moSupportedLanguages
        sPrompt = sPrompt & "• " & varLang & vbCrLf
    Next varLang
    
    sPrompt = sPrompt & vbCrLf & "Available commands:" & vbCrLf & _
                "• macro vba run [ARES]English" & vbCrLf & _
                "• macro vba run [ARES]Français"
    
    MsgBox sPrompt, vbInformation + vbOKOnly, "ARES Language Configuration"
    
    PromptForLanguageSelection = "" ' User must set manually
    Exit Function
    
ErrorHandler:
    ErrorHandler.HandleError Err.Description, Err.Number, Err.Source, "LangManager.PromptForLanguageSelection"
    PromptForLanguageSelection = ""
End Function