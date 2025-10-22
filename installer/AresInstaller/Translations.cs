using System.Collections.Generic;

namespace AresInstaller
{
    public static class Translations
    {
        private static Dictionary<string, Dictionary<string, string>> translations = new Dictionary<string, Dictionary<string, string>>()
        {
            // Window titles
            { "WindowTitle", new Dictionary<string, string> { { "EN", "ARES Installer" }, { "FR", "Installateur ARES" } } },
            { "ProductSelection", new Dictionary<string, string> { { "EN", "Bentley Product Selection" }, { "FR", "Sélection de produit Bentley" } } },
            
            // Buttons
            { "InstallButton", new Dictionary<string, string> { { "EN", "Install" }, { "FR", "Installer" } } },
            { "ExitButton", new Dictionary<string, string> { { "EN", "Exit" }, { "FR", "Quitter" } } },
            { "OKButton", new Dictionary<string, string> { { "EN", "OK" }, { "FR", "OK" } } },
            { "NextButton", new Dictionary<string, string> { { "EN", "Next" }, { "FR", "Suivant" } } },
            { "CancelButton", new Dictionary<string, string> { { "EN", "Cancel" }, { "FR", "Annuler" } } },

            // Bentley Product Selection
            { "SelectBentleyProduct", new Dictionary<string, string> { { "EN", "Select Bentley Product for ARES Integration" }, { "FR", "Sélectionnez un produit Bentley pour l'intégration ARES" } } },
            { "ConfigurationPath", new Dictionary<string, string> { { "EN", "Configuration Path:" }, { "FR", "Chemin de configuration :" } } },
            { "NoBentleyProducts", new Dictionary<string, string> { { "EN", "No Bentley products found in registry." }, { "FR", "Aucun produit Bentley trouvé dans le registre." } } },
            { "NoValidBentleyProducts", new Dictionary<string, string> { { "EN", "No valid Bentley products (MicroStation or MapPowerView) found." }, { "FR", "Aucun produit Bentley valide (MicroStation ou MapPowerView) trouvé." } } },
            { "NoConfigPath", new Dictionary<string, string> { { "EN", "This product does not have a configuration path defined." }, { "FR", "Ce produit n'a pas de chemin de configuration défini." } } },
            { "NoConfigPathAvailable", new Dictionary<string, string> { { "EN", "(No configuration path available)" }, { "FR", "(Aucun chemin de configuration disponible)" } } },
            { "FailedLoadProducts", new Dictionary<string, string> { { "EN", "Failed to load Bentley products: {0}" }, { "FR", "Échec du chargement des produits Bentley : {0}" } } },
            { "ProductSelectionError", new Dictionary<string, string> { { "EN", "Product Selection Error" }, { "FR", "Erreur de sélection de produit" } } },

            // Auto-configuration
            { "Configuration", new Dictionary<string, string> { { "EN", "Configuration" }, { "FR", "Configuration" } } },
            { "ConfigurationSuccess", new Dictionary<string, string> { { "EN", "ARES has been successfully configured for the selected product!" }, { "FR", "ARES a été configuré avec succès pour le produit sélectionné !" } } },
            { "ConfigurationError", new Dictionary<string, string> { { "EN", "Configuration error: {0}" }, { "FR", "Erreur de configuration : {0}" } } },
            { "BentleyProductPathNotFound", new Dictionary<string, string> { { "EN", "Bentley product path not found: {0}" }, { "FR", "Chemin du produit Bentley introuvable : {0}" } } },
            { "NoPersonalUcfFound", new Dictionary<string, string> { { "EN", "No Personal.ucf file found in: {0}" }, { "FR", "Aucun fichier Personal.ucf trouvé dans : {0}" } } },
            { "AlreadyConfigured", new Dictionary<string, string> { { "EN", "ARES is already configured for this product." }, { "FR", "ARES est déjà configuré pour ce produit." } } },
            { "ErrorSearchingUcfFiles", new Dictionary<string, string> { { "EN", "Error searching for UCF files: {0}" }, { "FR", "Erreur lors de la recherche des fichiers UCF : {0}" } } },
            { "ErrorProcessingUcfFile", new Dictionary<string, string> { { "EN", "Error processing file {0}: {1}" }, { "FR", "Erreur lors du traitement du fichier {0} : {1}" } } },
            
            // Status messages
            { "ReadyToInstall", new Dictionary<string, string> { { "EN", "Ready to install ARES" }, { "FR", "Prêt à installer ARES" } } },
            { "CheckingPrerequisites", new Dictionary<string, string> { { "EN", "Checking prerequisites..." }, { "FR", "Vérification des prérequis..." } } },
            { "CreatingDirectories", new Dictionary<string, string> { { "EN", "Creating installation directories..." }, { "FR", "Création des répertoires d'installation..." } } },
            { "Downloading", new Dictionary<string, string> { { "EN", "Downloading ARES components from GitHub..." }, { "FR", "Téléchargement des composants ARES depuis GitHub..." } } },
            { "Extracting", new Dictionary<string, string> { { "EN", "Extracting downloaded files..." }, { "FR", "Extraction des fichiers téléchargés..." } } },
            { "InstallingProject", new Dictionary<string, string> { { "EN", "Installing ARES project..." }, { "FR", "Installation du projet ARES..." } } },
            { "RegisteringCOM", new Dictionary<string, string> { { "EN", "Registering COM components..." }, { "FR", "Enregistrement des composants COM..." } } },
            { "InstallationCompleted", new Dictionary<string, string> { { "EN", "ARES installation completed successfully!" }, { "FR", "Installation d'ARES terminée avec succès !" } } },
            
            // Messages
            { "InstallSuccess", new Dictionary<string, string> { { "EN", "ARES installed successfully!" }, { "FR", "ARES installé avec succès !" } } },
            { "InstallError", new Dictionary<string, string> { { "EN", "Installation failed: {0}" }, { "FR", "Échec de l'installation : {0}" } } },
            { "InstallationComplete", new Dictionary<string, string> { { "EN", "Installation Complete" }, { "FR", "Installation terminée" } } },
            { "InstallationError", new Dictionary<string, string> { { "EN", "Installation Error" }, { "FR", "Erreur d'installation" } } },
            { "InstallationInProgress", new Dictionary<string, string> { { "EN", "Installation in progress. Please wait..." }, { "FR", "Installation en cours. Veuillez patienter..." } } },
            
            // Prerequisites
            { "PrerequisitesCheck", new Dictionary<string, string> { { "EN", "=== Prerequisites Check ===" }, { "FR", "=== Vérification des prérequis ===" } } },
            { "NotRunningAsAdmin", new Dictionary<string, string> { { "EN", "WARNING: Not running as Administrator - some features may not work" }, { "FR", "ATTENTION : Pas exécuté en tant qu'administrateur - certaines fonctionnalités peuvent ne pas fonctionner" } } },
            { "DotNetAvailable", new Dictionary<string, string> { { "EN", ".NET Framework 4.7.2+ available" }, { "FR", ".NET Framework 4.7.2+ disponible" } } },
            { "DotNetRequired", new Dictionary<string, string> { { "EN", ".NET Framework 4.7.2 or higher is required" }, { "FR", ".NET Framework 4.7.2 ou supérieur est requis" } } },
            
            // Directory Management
            { "CreatingDirs", new Dictionary<string, string> { { "EN", "=== Creating Directories ===" }, { "FR", "=== Création des répertoires ===" } } },
            { "Created", new Dictionary<string, string> { { "EN", "Created: {0}" }, { "FR", "Créé : {0}" } } },
            { "FailedCreateDirs", new Dictionary<string, string> { { "EN", "Failed to create directories: {0}" }, { "FR", "Échec de la création des répertoires : {0}" } } },
            
            // GitHub Download
            { "DownloadingFromGitHub", new Dictionary<string, string> { { "EN", "=== Downloading from GitHub ===" }, { "FR", "=== Téléchargement depuis GitHub ===" } } },
            { "FetchingReleaseInfoAPI", new Dictionary<string, string> { { "EN", "Fetching release information from GitHub API..." }, { "FR", "Récupération des informations de version depuis l'API GitHub..." } } },
            { "ParsingReleaseAssets", new Dictionary<string, string> { { "EN", "Parsing release assets..." }, { "FR", "Analyse des fichiers de la version..." } } },
            { "FetchingReleaseInfo", new Dictionary<string, string> { { "EN", "Fetching release information..." }, { "FR", "Récupération des informations de version..." } } },
            { "LatestVersion", new Dictionary<string, string> { { "EN", "Latest version: {0}" }, { "FR", "Dernière version : {0}" } } },
            { "NoAssetsFound", new Dictionary<string, string> { { "EN", "No assets found in the latest release. Please ensure files are uploaded to the GitHub release." }, { "FR", "Aucun fichier trouvé dans la dernière version. Assurez-vous que les fichiers sont téléchargés sur GitHub." } } },
            { "FoundAssets", new Dictionary<string, string> { { "EN", "Found {0} asset(s) to download" }, { "FR", "{0} fichier(s) à télécharger trouvé(s)" } } },
            { "FailedCreateDownloadDir", new Dictionary<string, string> { { "EN", "Failed to create download directory: {0}" }, { "FR", "Échec de la création du répertoire de téléchargement : {0}" } } },
            { "DownloadingFile", new Dictionary<string, string> { { "EN", "Downloading: {0}" }, { "FR", "Téléchargement : {0}" } } },
            { "DownloadComplete", new Dictionary<string, string> { { "EN", "Download complete. Total files in directory: {0}" }, { "FR", "Téléchargement terminé. Nombre total de fichiers : {0}" } } },
            { "FailedDownload", new Dictionary<string, string> { { "EN", "Failed to download from GitHub: {0}" }, { "FR", "Échec du téléchargement depuis GitHub : {0}" } } },
            { "FailedParseJSON", new Dictionary<string, string> { { "EN", "Failed to parse GitHub release JSON: {0}" }, { "FR", "Échec de l'analyse du JSON de version GitHub : {0}" } } },
            { "FailedDownloadAssets", new Dictionary<string, string> { { "EN", "Failed to download assets: {0}" }, { "FR", "Échec du téléchargement des fichiers : {0}" } } },
            
            // File Extraction
            { "ExtractingFiles", new Dictionary<string, string> { { "EN", "=== Extracting Downloaded Files ===" }, { "FR", "=== Extraction des fichiers téléchargés ===" } } },
            { "ExtractingZip", new Dictionary<string, string> { { "EN", "Extracting: {0}" }, { "FR", "Extraction : {0}" } } },
            { "Extracted", new Dictionary<string, string> { { "EN", "   Extracted: {0}" }, { "FR", "   Extrait : {0}" } } },
            { "Copied", new Dictionary<string, string> { { "EN", "Copied: {0}" }, { "FR", "Copié : {0}" } } },
            { "FailedExtract", new Dictionary<string, string> { { "EN", "Failed to extract files: {0}" }, { "FR", "Échec de l'extraction des fichiers : {0}" } } },
            
            // DLL Registration
            { "RegisteringComponents", new Dictionary<string, string> { { "EN", "=== Registering COM Components ===" }, { "FR", "=== Enregistrement des composants COM ===" } } },
            { "SearchingValidator", new Dictionary<string, string> { { "EN", "Searching for AresLicenseValidator DLL..." }, { "FR", "Recherche de AresLicenseValidator DLL..." } } },
            { "ValidatorNotFound", new Dictionary<string, string> { { "EN", "AresLicenseValidator.dll not found after copying to Rsc folder" }, { "FR", "AresLicenseValidator.dll introuvable après copie dans le dossier Rsc" } } },
            { "ValidatorNotFoundAtPath", new Dictionary<string, string> { { "EN", "AresLicenseValidator.dll not found at path: {0}" }, { "FR", "AresLicenseValidator.dll introuvable au chemin : {0}" } } },
            { "FoundDLL", new Dictionary<string, string> { { "EN", "Found DLL: {0}" }, { "FR", "DLL trouvée : {0}" } } },
            { "COMRegistrationComplete", new Dictionary<string, string> { { "EN", "COM registration completed" }, { "FR", "Enregistrement COM terminé" } } },
            { "ProcessingDLL", new Dictionary<string, string> { { "EN", "Processing: {0}" }, { "FR", "Traitement : {0}" } } },
            { "SameVersionInstalled", new Dictionary<string, string> { { "EN", "  Same version already installed - Skipping" }, { "FR", "  Même version déjà installée - Ignoré" } } },
            { "BackedUpOldVersion", new Dictionary<string, string> { { "EN", "  Backed up old version to: {0}" }, { "FR", "  Ancienne version sauvegardée dans : {0}" } } },
            { "BackedUpOldTLB", new Dictionary<string, string> { { "EN", "  Backed up old TLB to: {0}" }, { "FR", "  Ancien TLB sauvegardé dans : {0}" } } },
            { "CopiedDLL", new Dictionary<string, string> { { "EN", "  ✓ Copied: {0} ({1} bytes)" }, { "FR", "  ✓ Copié : {0} ({1} octets)" } } },
            { "FailedCopyDLL", new Dictionary<string, string> { { "EN", "Failed to copy DLL to: {0}" }, { "FR", "Échec de la copie de la DLL vers : {0}" } } },
            { "CopiedTLB", new Dictionary<string, string> { { "EN", "  ✓ Copied TLB: {0}" }, { "FR", "  ✓ TLB copié : {0}" } } },
            { "RegisteringDLL", new Dictionary<string, string> { { "EN", "Registering: {0}" }, { "FR", "Enregistrement : {0}" } } },
            { "RegAsmNotFound", new Dictionary<string, string> { { "EN", "RegAsm.exe not found. Please install .NET Framework Developer Pack." }, { "FR", "RegAsm.exe introuvable. Veuillez installer le pack développeur .NET Framework." } } },
            { "DLLRegisteredSuccess", new Dictionary<string, string> { { "EN", "DLL registered successfully" }, { "FR", "DLL enregistrée avec succès" } } },
            { "RegAsmFailed", new Dictionary<string, string> { { "EN", "RegAsm failed (Exit code: {0}): {1}" }, { "FR", "Échec de RegAsm (Code de sortie : {0}) : {1}" } } },
            { "FailedRegisterDLLs", new Dictionary<string, string> { { "EN", "Failed to register DLLs: {0}" }, { "FR", "Échec de l'enregistrement des DLL : {0}" } } },
            { "FailedExecuteRegAsm", new Dictionary<string, string> { { "EN", "Failed to execute RegAsm: {0}" }, { "FR", "Échec de l'exécution de RegAsm : {0}" } } },
            
            // Project Installation
            { "InstallingARES", new Dictionary<string, string> { { "EN", "=== Installing ARES Project ===" }, { "FR", "=== Installation du projet ARES ===" } } },
            { "CopiedMVBA", new Dictionary<string, string> { { "EN", "Copied ARES.mvba to: {0}" }, { "FR", "ARES.mvba copié dans : {0}" } } },
            { "MVBANotFound", new Dictionary<string, string> { { "EN", "ARES.mvba not found in download" }, { "FR", "ARES.mvba introuvable dans le téléchargement" } } },
            { "FailedCopyProject", new Dictionary<string, string> { { "EN", "Failed to copy ARES project: {0}" }, { "FR", "Échec de la copie du projet ARES : {0}" } } },
            
            // Cleanup
            { "CleaningUp", new Dictionary<string, string> { { "EN", "=== Cleaning Up ===" }, { "FR", "=== Nettoyage ===" } } },
            { "CleanupCompleted", new Dictionary<string, string> { { "EN", "Cleanup completed" }, { "FR", "Nettoyage terminé" } } },
            { "CleanupWarning", new Dictionary<string, string> { { "EN", "Cleanup warning: {0}" }, { "FR", "Avertissement de nettoyage : {0}" } } },
            
            // Installation Summary
            { "InstallationSummary", new Dictionary<string, string> { { "EN", "=== Installation Summary ===" }, { "FR", "=== Résumé de l'installation ===" } } },
            { "MainProject", new Dictionary<string, string> { { "EN", "Main project: {0}" }, { "FR", "Projet principal : {0}" } } },
            { "DLLComponents", new Dictionary<string, string> { { "EN", "DLL components: {0}" }, { "FR", "Composants DLL : {0}" } } },
            { "COMRegistered", new Dictionary<string, string> { { "EN", "COM components registered" }, { "FR", "Composants COM enregistrés" } } },
            { "NextSteps", new Dictionary<string, string> { { "EN", "Next steps:" }, { "FR", "Prochaines étapes :" } } },
            { "Step1", new Dictionary<string, string> { { "EN", "1. Open MicroStation" }, { "FR", "1. Ouvrir MicroStation" } } },
            { "Step2", new Dictionary<string, string> { { "EN", "2. Load ARES.mvba from the MicroStation VBA Manager" }, { "FR", "2. Charger ARES.mvba depuis le gestionnaire VBA de MicroStation" } } },
            { "Step3", new Dictionary<string, string> { { "EN", "3. Create license" }, { "FR", "3. Créer une licence" } } },
        };

        public static string Get(string key, string language)
        {
            if (translations.ContainsKey(key) && translations[key].ContainsKey(language))
            {
                return translations[key][language];
            }
            return key; // Return key if translation not found
        }

        public static string Format(string key, string language, params object[] args)
        {
            string template = Get(key, language);
            return string.Format(template, args);
        }
    }
}