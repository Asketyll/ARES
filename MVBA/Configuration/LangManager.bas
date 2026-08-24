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
    moTranslations.Add "EN_ZoningNoBufferCreated", "No buffer could be created for any of the {0} element(s) found."
    moTranslations.Add "EN_ZoningSomeBuffersFailed", "{0} of {1} element(s) could not be buffered and were skipped."
    moTranslations.Add "EN_ZoneExportGUIOptionsCaption", "Edit zone export options:"
    moTranslations.Add "EN_ZoneExportGUIOptionsEdit_Level_Region_CommandCaption", "Edit zone level"
    moTranslations.Add "EN_ZoneExportGUIOptionsEdit_Level_Candidate_CommandCaption", "Edit filter level"
    moTranslations.Add "EN_ZoneExportGUIOptionsRound_LabelCaption", "Decimal places:"
    moTranslations.Add "EN_ZoneExportGUIOptionsUse_Dialog_LabelCaption", "Choose file location"
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
    moTranslations.Add "EN_FormDefaultsRestoreConfirm", "Restore every option on this form to its default value?"
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
    moTranslations.Add "EN_ZoneExportGroupByID", "(none)"
    moTranslations.Add "EN_ZoneExportGUIOptionsPerZone_LabelCaption", "Break down per zone"
    moTranslations.Add "EN_ZoneExportGUIOptionsZoneProperty_LabelCaption", "Naming property:"
    moTranslations.Add "EN_ZoneExportGUIOptionsOpenAfter_LabelCaption", "Open once exported"
    ' --- Story 8-2: restore-defaults tooltip + element-picker OK/Cancel ---
    moTranslations.Add "EN_FormResetDefaultsTip", "Reset every option on this panel to its default value."
    moTranslations.Add "EN_FormHelpTip", "Open the online wiki page for this feature's full syntax reference."
    ' Tooltips (ControlTipText) - Property Tagging
    moTranslations.Add "EN_PropertyTaggingGUIOptionsMain_LabelTip", "Attach ARES custom properties automatically when elements are created or modified."
    moTranslations.Add "EN_PropertyTaggingGUIOptionsEditRules_CommandTip", "Rules: Lvl[level], Cell[name], Type[type]; & = AND; ! negates; * / ? wildcards; a leading @ attaches to the OTHER group members. Example: Type[Cell]&!Cell[A]=Repere ; @Cell[ETI0*]=Repere"
    moTranslations.Add "EN_PropertyRuleInvalid", "ARES: Property rule invalid, not saved (expected [@]Lvl/Cell/Type[name]&...=prop|prop)"
    ' --- Property Calculation options GUI + statuses ---
    moTranslations.Add "EN_CalculationGUIOptionsCaption", "Property Calculation options"
    moTranslations.Add "EN_CalculationGUIOptionsMain_LabelCaption", "Enable property calculation"
    moTranslations.Add "EN_CalculationGUIOptionsMain_LabelTip", "When on, properties get their values from the calculation rules below (Prop[name]=Source): a group label cell's text (CellText) or coordinates (CellCoord) or ID (CellId), a fixed value (Value), the element's own coordinates (Coord) or its own ID (Id). Values are only written where the property is already attached - attaching is Property Tagging's job."
    moTranslations.Add "EN_CalculationGUIOptionsDetachEmpty_LabelCaption", "Remove emptied properties after calculation"
    moTranslations.Add "EN_CalculationGUIOptionsDetachEmpty_LabelTip", "When on, a property whose value is emptied by the calculation is detached (the tagger removes it) instead of kept empty; a rule that still mandates the property re-attaches it empty."
    moTranslations.Add "EN_CalculationValueRejected", "ARES: Property calculation - value rejected by the target property"
    moTranslations.Add "EN_CalculationNoTarget", "ARES: Property calculation - no group member carries the target property; enable Property Tagging and check the property is attached (DGNLib)"
    moTranslations.Add "EN_CalculationMultipleTriggers", "ARES: Property calculation - several trigger cells in this group; the last-modified one sets the value"
    moTranslations.Add "EN_CalculationMultipleLvlTriggers", "ARES: Property calculation - several level-matching trigger elements in this group; the last-modified one sets the value"
    moTranslations.Add "EN_CalculationMultipleGeometries", "ARES: Property calculation - several measurable geometries in this group; the first one found sets the length"
    moTranslations.Add "EN_CalcRuleInvalid", "ARES: Calc rule invalid, not saved (expected Prop[name][&...]=Source; Source = CellText|CellCoord|CellId|CellLvl|CellColor|CellStyle|CellWeight|LvlColor|LvlStyle|LvlWeight[pattern] / Value[text] / Coord|Length|GroupLength[n] / Id|Lvl|Color|Style|Weight)"
    moTranslations.Add "EN_CalculationGUIOptionsCalcRules_Tip", "Calc rules: Prop[name] [&Lvl/Cell/Type[..]] = CellText|CellCoord|CellId|CellLvl|CellColor|CellStyle|CellWeight|LvlColor|LvlStyle|LvlWeight[pattern] | Value[text] | Coord|Length|GroupLength[n] | Id|Lvl|Color|Style|Weight. pattern may use '|' for several name alternatives, e.g. ASUF*|SP0*. First matching rule per property wins (specific rules first). Example: Prop[Repere]&Cell[ETIREF]=Value[REF] ; Prop[Repere]=CellText[ETI*] ; Prop[XY]=Coord ; Prop[Coordonnee]=CellCoord[ASUF*|SP0*] ; Prop[Longueur]=GroupLength[1]"
    ' --- Custom-property DGNLib round trip (OpenPropertyLibrary key-in) ---
    moTranslations.Add "EN_PropertyLibraryNotFound", "ARES: Custom properties - the ARES DGNLib was not found (check MS_DGNLIBLIST)"
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
    moTranslations.Add "EN_ZoneExportGUIOptionsPerZone_LabelTip", "Splits each length by zone; names each zone with the chosen property."
    moTranslations.Add "EN_ZoneExportGUIOptionsZoneProperty_LabelTip", "Custom property used as the zone name. Empty or invalid = zones numbered Zone 1, Zone 2, ..."
    moTranslations.Add "EN_ZoneExportGUIOptionsRound_LabelTip", "Decimal places for exported lengths (0-10)."
    moTranslations.Add "EN_ZoneExportGUIOptionsUse_Dialog_LabelTip", "When on, prompt for the save location; otherwise auto-name the file."
    moTranslations.Add "EN_ZoneExportGUIOptionsOpenAfter_LabelTip", "When on, the exported file opens in Excel after the export."
    ' Tooltips - Auto Lengths
    ' --- Property Rendering (epic 15) ---
    moTranslations.Add "EN_RenderTokenUnknown", "ARES: Property rendering - unknown property in a Prop[...] token; the token is left as plain text"
    moTranslations.Add "EN_RenderValueUnsupported", "ARES: Property rendering - the property value is not text; the token is left as plain text"
    moTranslations.Add "EN_RenderValueIllegalChars", "ARES: Property rendering - the value contains a reserved character or a line break; nothing was written"
    moTranslations.Add "EN_RenderMetadataInvalid", "ARES: Property rendering - the stored binding is inconsistent with its text; nothing was rendered"
    moTranslations.Add "EN_RenderMetadataUnreadable", "ARES: Property rendering - the stored binding could not be read or saved; nothing was changed"
    moTranslations.Add "EN_RenderSchemaUnsupported", "ARES: Property rendering - this binding was written by a newer ARES; it is left untouched"
    moTranslations.Add "EN_RenderLibraryMissing", "ARES: Property rendering - internal item type library missing; update the ARES resources (DGNLib)"
    moTranslations.Add "EN_RenderAmbiguousEdit", "ARES: Property rendering - the edited text could not be matched; only the visible Prop[...] tokens stay linked"
    moTranslations.Add "EN_RenderBindingReleased", "ARES: Property rendering - the edited text carries no Prop[...] token any more; the link was removed. Re-type a token, then run BindPropertyRender"
    moTranslations.Add "EN_RenderElementLocked", "ARES: Property rendering - the element is locked; nothing was rendered"
    moTranslations.Add "EN_RenderSubIdDrift", "ARES: Property rendering - the linked text could not be located in this cell; nothing was written"
    moTranslations.Add "EN_RenderDuplicateToken", "ARES: Property rendering - the same property is used twice in one text; use it only once"
    moTranslations.Add "EN_RenderAdjacentTokens", "ARES: Property rendering - two tokens must be separated by some text"
    moTranslations.Add "EN_RenderTextNodeInCellRefused", "ARES: Property rendering - a multi-line text inside a cell used by a calc rule cannot be linked"
    moTranslations.Add "EN_RenderNotBound", "ARES: Property rendering - the property is not attached to this element; attach it, then run BindPropertyRender"
    moTranslations.Add "EN_RenderBindDone", "ARES: Property rendering - text linked"
    moTranslations.Add "EN_RenderCycleWarning", "ARES: Property rendering - this value is computed from the text of this very cell; the rendered text is excluded from that computation"
    moTranslations.Add "EN_RenderNoSelection", "ARES: Property rendering - select the text(s) to link first"
    moTranslations.Add "EN_RenderDisabled", "ARES: Property rendering is disabled (ARES_Text_Render)"
    moTranslations.Add "EN_RenderValueGoverned", "ARES: Property rendering - the value of {0} is set by a calc rule; editing the text will not change it"

    ' --- Property Actuator (epic 16) ---
    moTranslations.Add "EN_ActuatorColorInvalid", "ARES: Property actuator - the pilot property's value is not a valid color index"
    moTranslations.Add "EN_ActuatorLevelInvalid", "ARES: Property actuator - the pilot property's value does not name an existing, unlocked level"
    moTranslations.Add "EN_ActuatorSelfRatchetRefused", "ARES: Property actuator - refused: the pilot property is itself computed from this same attribute"

    ' Property Rendering options form (epic 15). The colour-sync and ATLAS label settings moved here from
    ' the Auto Lengths options form: all three serve DISPLAY, and CellRedreaw - the sole consumer of the two
    ' ATLAS settings - is called by the renderer's only text write.
    moTranslations.Add "EN_RenderingGUIOptionsCaption", "Property Rendering options"
    moTranslations.Add "EN_RenderingGUIOptionsMain_LabelCaption", "Enable property rendering"
    moTranslations.Add "EN_RenderingGUIOptionsMain_LabelTip", "Write a linked property's value into a text carrying a Prop[...] token (ARES_Text_Render)."
    moTranslations.Add "EN_RenderingGUIOptionsColor_LabelCaption", "Sync text colour with the linked geometry"
    moTranslations.Add "EN_RenderingGUIOptionsColor_LabelTip", "A text linked to a geometry takes that geometry's colour when the colour changes (ARES_Only_Color_Update)."
    moTranslations.Add "EN_RenderingGUIOptionsCell_LabelCaption", "Rebuild ATLAS label cells after a text change"
    moTranslations.Add "EN_RenderingGUIOptionsCell_LabelTip", "Rebuild the leader-label cell geometry once its text has changed, so the frame follows the new text (ARES_Update_ATLASCellLabel)."
    moTranslations.Add "EN_RenderingGUIOptionsEdit_Cells_List_CommandCaption", "Edit ATLAS cell names"
    moTranslations.Add "EN_RenderingGUIOptionsEdit_Cells_List_CommandTip", "Cell names treated as ATLAS labels, separated by | (ARES_Cell_Is_Label_Name). Enter commits, Esc cancels."
    moTranslations.Add "EN_RenderingGUIOptionsActuatorSection_LabelCaption", "Property Actuator (paints Color/Level)"
    moTranslations.Add "EN_RenderingGUIOptionsActuateColor_LabelCaption", "Paint element Color from ARES_Color"
    moTranslations.Add "EN_RenderingGUIOptionsActuateColor_LabelTip", "An element's own Color follows its ARES_Color property, whenever they differ (ARES_Actuate_Color). Attach ARES_Color to your elements via a PropertyTagging rule (e.g. Lvl[WALLS]=Commune|ARES_Color)."
    moTranslations.Add "EN_RenderingGUIOptionsActuateLevel_LabelCaption", "Paint element Level from ARES_Lvl"
    moTranslations.Add "EN_RenderingGUIOptionsActuateLevel_LabelTip", "An element's own Level follows its ARES_Lvl property, whenever they differ (ARES_Actuate_Level). Attach ARES_Lvl to your elements via a PropertyTagging rule (e.g. Lvl[WALLS]=Commune|ARES_Lvl)."

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
    moTranslations.Add "FR_ZoningNoBufferCreated", "Aucun buffer n'a pu être créé pour les {0} élément(s) trouvé(s)."
    moTranslations.Add "FR_ZoningSomeBuffersFailed", "{0} des {1} élément(s) n'ont pas pu être bufférisés et ont été ignorés."
    moTranslations.Add "FR_ZoneExportGUIOptionsCaption", "Modifier les options d'export de zone :"
    moTranslations.Add "FR_ZoneExportGUIOptionsEdit_Level_Region_CommandCaption", "Modifier le niveau de zone"
    moTranslations.Add "FR_ZoneExportGUIOptionsEdit_Level_Candidate_CommandCaption", "Modifier le niveau de filtre"
    moTranslations.Add "FR_ZoneExportGUIOptionsRound_LabelCaption", "Décimales :"
    moTranslations.Add "FR_ZoneExportGUIOptionsUse_Dialog_LabelCaption", "Choisir l'emplacement du fichier"
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
    moTranslations.Add "FR_FormDefaultsRestoreConfirm", "Réinitialiser toutes les options de ce formulaire à leur valeur par défaut ?"
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
    moTranslations.Add "FR_ZoneExportGroupByID", "vide"
    moTranslations.Add "FR_ZoneExportGUIOptionsPerZone_LabelCaption", "Répartir par zone"
    moTranslations.Add "FR_ZoneExportGUIOptionsZoneProperty_LabelCaption", "Propriété de nommage :"
    moTranslations.Add "FR_ZoneExportGUIOptionsOpenAfter_LabelCaption", "Ouvrir une fois exportée"
    ' --- Story 8-2 : info-bulle reinitialisation + OK/Annuler du selecteur d'elements ---
    moTranslations.Add "FR_FormResetDefaultsTip", "Réinitialise chaque option de ce panneau à sa valeur par défaut."
    moTranslations.Add "FR_FormHelpTip", "Ouvre la page wiki en ligne de référence complète pour cette fonctionnalité."
    ' Tooltips (ControlTipText) - Property Tagging
    moTranslations.Add "FR_PropertyTaggingGUIOptionsMain_LabelTip", "Attache automatiquement les propriétés ARES à la création ou à la modification d'éléments."
    moTranslations.Add "FR_PropertyTaggingGUIOptionsEditRules_CommandTip", "Règles : Lvl[niveau], Cell[nom], Type[type] ; & = ET ; ! nie ; * / ? jokers ; un @ en tête attache aux AUTRES membres du groupe. Exemple : Type[Cell]&!Cell[A]=Repere ; @Cell[ETI0*]=Repere"
    moTranslations.Add "FR_PropertyRuleInvalid", "ARES : Règle de propriété invalide, non enregistrée (attendu [@]Lvl/Cell/Type[nom]&...=prop|prop)"
    ' --- Property Calculation options GUI + statuses ---
    moTranslations.Add "FR_CalculationGUIOptionsCaption", "Options de calcul de propriété"
    moTranslations.Add "FR_CalculationGUIOptionsMain_LabelCaption", "Activer le calcul de propriété"
    moTranslations.Add "FR_CalculationGUIOptionsMain_LabelTip", "Si activé, les propriétés reçoivent leur valeur selon les règles de calcul ci-dessous (Prop[nom]=Source) : le texte d'une cellule étiquette du groupe (CellText) ou ses coordonnées (CellCoord) ou son ID (CellId), une valeur fixe (Value), les coordonnées propres de l'élément (Coord) ou son propre ID (Id). Les valeurs ne sont écrites que là où la propriété est déjà attachée - l'attache relève de l'étiquetage de propriété."
    moTranslations.Add "FR_CalculationGUIOptionsDetachEmpty_LabelCaption", "Supprimer les propriétés vidées après calcul"
    moTranslations.Add "FR_CalculationGUIOptionsDetachEmpty_LabelTip", "Si activé, une propriété dont la valeur est vidée par le calcul est détachée (le tagueur la retire) au lieu d'être conservée vide ; une règle qui impose encore la propriété la ré-attache vide."
    moTranslations.Add "FR_CalculationValueRejected", "ARES : Calcul de propriété - valeur refusée par la propriété cible"
    moTranslations.Add "FR_CalculationNoTarget", "ARES : Calcul de propriété - aucun membre du groupe ne porte la propriété cible ; activez l'étiquetage de propriété et vérifiez que la propriété est attachée (DGNLib)"
    moTranslations.Add "FR_CalculationMultipleTriggers", "ARES : Calcul de propriété - plusieurs cellules déclencheuses dans ce groupe ; la dernière modifiée impose la valeur"
    moTranslations.Add "FR_CalculationMultipleLvlTriggers", "ARES : Calcul de propriété - plusieurs éléments déclencheurs (niveau correspondant) dans ce groupe ; le dernier modifié impose la valeur"
    moTranslations.Add "FR_CalculationMultipleGeometries", "ARES : Calcul de propriété - plusieurs géométries mesurables dans ce groupe ; la première trouvée impose la longueur"
    moTranslations.Add "FR_CalcRuleInvalid", "ARES : Règle de calcul invalide, non enregistrée (attendu Prop[nom][&...]=Source ; Source = CellText|CellCoord|CellId|CellLvl|CellColor|CellStyle|CellWeight|LvlColor|LvlStyle|LvlWeight[motif] / Value[texte] / Coord|Length|GroupLength[n] / Id|Lvl|Color|Style|Weight)"
    moTranslations.Add "FR_CalculationGUIOptionsCalcRules_Tip", "Règles de calcul : Prop[nom] [&Lvl/Cell/Type[..]] = CellText|CellCoord|CellId|CellLvl|CellColor|CellStyle|CellWeight|LvlColor|LvlStyle|LvlWeight[motif] | Value[texte] | Coord|Length|GroupLength[n] | Id|Lvl|Color|Style|Weight. Le motif peut utiliser '|' pour plusieurs alternatives de nom, ex. ASUF*|SP0*. La première règle qui correspond par propriété gagne (règles spécifiques d'abord). Exemple : Prop[Repere]&Cell[ETIREF]=Value[REF] ; Prop[Repere]=CellText[ETI*] ; Prop[XY]=Coord ; Prop[Coordonnee]=CellCoord[ASUF*|SP0*] ; Prop[Longueur]=GroupLength[1]"
    ' --- Custom-property DGNLib round trip (OpenPropertyLibrary key-in) ---
    moTranslations.Add "FR_PropertyLibraryNotFound", "ARES : Propriétés personnalisées - DGNLib ARES introuvable (vérifiez MS_DGNLIBLIST)"
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
    moTranslations.Add "FR_ZoneExportGUIOptionsPerZone_LabelTip", "Répartit chaque longueur par zone, nomme chaque zone via la propriété choisie."
    moTranslations.Add "FR_ZoneExportGUIOptionsZoneProperty_LabelTip", "Propriété personnalisée servant de nom de zone. Vide ou invalide = zones numérotées Zone 1, Zone 2, ..."
    moTranslations.Add "FR_ZoneExportGUIOptionsRound_LabelTip", "Décimales pour les longueurs exportées (0-10)."
    moTranslations.Add "FR_ZoneExportGUIOptionsUse_Dialog_LabelTip", "Si activé, demande l'emplacement d'export ; sinon nomme le fichier automatiquement."
    moTranslations.Add "FR_ZoneExportGUIOptionsOpenAfter_LabelTip", "Si activé, le fichier exporté s'ouvre dans Excel après l'export."
    ' Tooltips - Auto Lengths
    ' --- Property Rendering (epic 15) ---
    moTranslations.Add "FR_RenderTokenUnknown", "ARES : Rendu de propriété - propriété inconnue dans un jeton Prop[...] ; le jeton reste du texte brut"
    moTranslations.Add "FR_RenderValueUnsupported", "ARES : Rendu de propriété - la valeur de la propriété n'est pas du texte ; le jeton reste du texte brut"
    moTranslations.Add "FR_RenderValueIllegalChars", "ARES : Rendu de propriété - la valeur contient un caractère réservé ou un saut de ligne ; rien n'a été écrit"
    moTranslations.Add "FR_RenderMetadataInvalid", "ARES : Rendu de propriété - la liaison enregistrée est incohérente avec son texte ; rien n'a été rendu"
    moTranslations.Add "FR_RenderMetadataUnreadable", "ARES : Rendu de propriété - la liaison enregistrée n'a pu être lue ou enregistrée ; rien n'a été modifié"
    moTranslations.Add "FR_RenderSchemaUnsupported", "ARES : Rendu de propriété - cette liaison a été écrite par un ARES plus récent ; elle est laissée intacte"
    moTranslations.Add "FR_RenderLibraryMissing", "ARES : Rendu de propriété - bibliothèque de types interne absente ; mettez à jour les ressources ARES (DGNLib)"
    moTranslations.Add "FR_RenderAmbiguousEdit", "ARES : Rendu de propriété - le texte modifié n'a pu être reconnu ; seuls les jetons Prop[...] visibles restent liés"
    moTranslations.Add "FR_RenderBindingReleased", "ARES : Rendu de propriété - le texte modifié ne comporte plus aucun jeton Prop[...] ; la liaison a été supprimée. Retapez un jeton puis lancez BindPropertyRender"
    moTranslations.Add "FR_RenderElementLocked", "ARES : Rendu de propriété - l'élément est verrouillé ; rien n'a été rendu"
    moTranslations.Add "FR_RenderSubIdDrift", "ARES : Rendu de propriété - le texte lié est introuvable dans cette cellule ; rien n'a été écrit"
    moTranslations.Add "FR_RenderDuplicateToken", "ARES : Rendu de propriété - la même propriété est utilisée deux fois dans un texte ; ne l'utilisez qu'une fois"
    moTranslations.Add "FR_RenderAdjacentTokens", "ARES : Rendu de propriété - deux jetons doivent être séparés par du texte"
    moTranslations.Add "FR_RenderTextNodeInCellRefused", "ARES : Rendu de propriété - un texte multiligne dans une cellule utilisée par une règle de calcul ne peut pas être lié"
    moTranslations.Add "FR_RenderNotBound", "ARES : Rendu de propriété - la propriété n'est pas attachée à cet élément ; attachez-la puis lancez BindPropertyRender"
    moTranslations.Add "FR_RenderBindDone", "ARES : Rendu de propriété - texte lié"
    moTranslations.Add "FR_RenderCycleWarning", "ARES : Rendu de propriété - cette valeur est calculée à partir du texte de cette même cellule ; le texte rendu est exclu de ce calcul"
    moTranslations.Add "FR_RenderNoSelection", "ARES : Rendu de propriété - sélectionnez d'abord le ou les textes à lier"
    moTranslations.Add "FR_RenderDisabled", "ARES : Le rendu de propriété est désactivé (ARES_Text_Render)"
    moTranslations.Add "FR_RenderValueGoverned", "ARES : Rendu de propriété - la valeur de {0} est imposée par une règle de calcul ; modifier le texte ne la changera pas"

    ' --- Actionneur de propriété (epic 16) ---
    moTranslations.Add "FR_ActuatorColorInvalid", "ARES : Actionneur de propriété - la valeur de la propriété pilote n'est pas un indice de couleur valide"
    moTranslations.Add "FR_ActuatorLevelInvalid", "ARES : Actionneur de propriété - la valeur de la propriété pilote ne désigne pas un niveau existant et non verrouillé"
    moTranslations.Add "FR_ActuatorSelfRatchetRefused", "ARES : Actionneur de propriété - refusé : la propriété pilote est elle-même calculée à partir de ce même attribut"

    ' Formulaire d'options du Rendu de propriétés (epic 15)
    moTranslations.Add "FR_RenderingGUIOptionsCaption", "Options du Rendu de propriétés"
    moTranslations.Add "FR_RenderingGUIOptionsMain_LabelCaption", "Activer le rendu de propriétés"
    moTranslations.Add "FR_RenderingGUIOptionsMain_LabelTip", "Écrit la valeur d'une propriété liée dans un texte portant un token Prop[...] (ARES_Text_Render)."
    moTranslations.Add "FR_RenderingGUIOptionsColor_LabelCaption", "Synchroniser la couleur du texte avec la géométrie liée"
    moTranslations.Add "FR_RenderingGUIOptionsColor_LabelTip", "Un texte lié à une géométrie prend la couleur de celle-ci quand elle change (ARES_Only_Color_Update)."
    moTranslations.Add "FR_RenderingGUIOptionsCell_LabelCaption", "Reconstruire les cellules d'étiquette ATLAS après modification du texte"
    moTranslations.Add "FR_RenderingGUIOptionsCell_LabelTip", "Reconstruit la géométrie de la cellule d'étiquette une fois son texte modifié, pour que le cadre suive le nouveau texte (ARES_Update_ATLASCellLabel)."
    moTranslations.Add "FR_RenderingGUIOptionsEdit_Cells_List_CommandCaption", "Modifier les noms de cellules ATLAS"
    moTranslations.Add "FR_RenderingGUIOptionsEdit_Cells_List_CommandTip", "Noms de cellules traitées comme étiquettes ATLAS, séparés par | (ARES_Cell_Is_Label_Name). Entrée valide, Échap annule."
    moTranslations.Add "FR_RenderingGUIOptionsActuatorSection_LabelCaption", "Actionneur de propriété (peint couleur/niveau)"
    moTranslations.Add "FR_RenderingGUIOptionsActuateColor_LabelCaption", "Peindre la couleur de l'élément depuis ARES_Color"
    moTranslations.Add "FR_RenderingGUIOptionsActuateColor_LabelTip", "La couleur propre de l'élément suit sa propriété ARES_Color, dès qu'elles diffèrent (ARES_Actuate_Color). Attachez ARES_Color à vos éléments via une règle PropertyTagging (ex. Lvl[WALLS]=Commune|ARES_Color)."
    moTranslations.Add "FR_RenderingGUIOptionsActuateLevel_LabelCaption", "Peindre le niveau de l'élément depuis ARES_Lvl"
    moTranslations.Add "FR_RenderingGUIOptionsActuateLevel_LabelTip", "Le niveau propre de l'élément suit sa propriété ARES_Lvl, dès qu'ils diffèrent (ARES_Actuate_Level). Attachez ARES_Lvl à vos éléments via une règle PropertyTagging (ex. Lvl[WALLS]=Commune|ARES_Lvl)."

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