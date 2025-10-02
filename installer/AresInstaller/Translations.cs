using System.Collections.Generic;

namespace AresInstaller
{
    public static class Translations
    {
        private static Dictionary<string, Dictionary<string, string>> translations = new Dictionary<string, Dictionary<string, string>>()
        {
            // Window titles
            { "WindowTitle", new Dictionary<string, string> { { "EN", "ARES Installer" }, { "FR", "Installateur ARES" } } },
            
            // Buttons
            { "InstallButton", new Dictionary<string, string> { { "EN", "Install" }, { "FR", "Installer" } } },
            { "ExitButton", new Dictionary<string, string> { { "EN", "Exit" }, { "FR", "Quitter" } } },

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
            
            // Prerequisites
            { "RunningAsAdmin", new Dictionary<string, string> { { "EN", "Running as Administrator" }, { "FR", "Exécution en tant qu'administrateur" } } },
            { "NotRunningAsAdmin", new Dictionary<string, string> { { "EN", "WARNING: Not running as Administrator - some features may not work" }, { "FR", "ATTENTION : Pas exécuté en tant qu'administrateur - certaines fonctionnalités peuvent ne pas fonctionner" } } },
            { "DotNetAvailable", new Dictionary<string, string> { { "EN", ".NET Framework 4.7.2+ available" }, { "FR", ".NET Framework 4.7.2+ disponible" } } },
            
            // Logs
            { "PrerequisitesCheck", new Dictionary<string, string> { { "EN", "=== Prerequisites Check ===" }, { "FR", "=== Vérification des prérequis ===" } } },
            { "CreatingDirs", new Dictionary<string, string> { { "EN", "=== Creating Directories ===" }, { "FR", "=== Création des répertoires ===" } } },
            { "DownloadingFromGitHub", new Dictionary<string, string> { { "EN", "=== Downloading from GitHub ===" }, { "FR", "=== Téléchargement depuis GitHub ===" } } },
            { "ExtractingFiles", new Dictionary<string, string> { { "EN", "=== Extracting Downloaded Files ===" }, { "FR", "=== Extraction des fichiers téléchargés ===" } } },
            { "InstallingARES", new Dictionary<string, string> { { "EN", "=== Installing ARES Project ===" }, { "FR", "=== Installation du projet ARES ===" } } },
            { "RegisteringComponents", new Dictionary<string, string> { { "EN", "=== Registering COM Components ===" }, { "FR", "=== Enregistrement des composants COM ===" } } },
            { "InstallationSummary", new Dictionary<string, string> { { "EN", "=== Installation Summary ===" }, { "FR", "=== Résumé de l'installation ===" } } },
            { "CleaningUp", new Dictionary<string, string> { { "EN", "=== Cleaning Up ===" }, { "FR", "=== Nettoyage ===" } } },

            { "Created", new Dictionary<string, string> { { "EN", "Created: {0}" }, { "FR", "Créé : {0}" } } },
            { "Copied", new Dictionary<string, string> { { "EN", "Copied: {0}" }, { "FR", "Copié : {0}" } } },
            { "Downloaded", new Dictionary<string, string> { { "EN", "Downloaded: {0} ({1} KB)" }, { "FR", "Téléchargé : {0} ({1} KB)" } } },
            { "Extracted", new Dictionary<string, string> { { "EN", "Extracted: {0}" }, { "FR", "Extrait : {0}" } } },
            
            // Summary
            { "MainProject", new Dictionary<string, string> { { "EN", "Main project: {0}" }, { "FR", "Projet principal : {0}" } } },
            { "DLLComponents", new Dictionary<string, string> { { "EN", "DLL components: {0}" }, { "FR", "Composants DLL : {0}" } } },
            { "COMRegistered", new Dictionary<string, string> { { "EN", "COM components registered" }, { "FR", "Composants COM enregistrés" } } },
            { "LicenseTools", new Dictionary<string, string> { { "EN", "License tools: {0}" }, { "FR", "Outils de licence : {0}" } } },
            { "NextSteps", new Dictionary<string, string> { { "EN", "Next steps:" }, { "FR", "Prochaines étapes :" } } },
            { "Step1", new Dictionary<string, string> { { "EN", "1. Generate licenses using Tools/Generate-License.bat" }, { "FR", "1. Générer des licences avec Tools/Generate-License.bat" } } },
            { "Step2", new Dictionary<string, string> { { "EN", "2. Place license files on network share" }, { "FR", "2. Placer les fichiers de licence sur le partage réseau" } } },
            { "Step3", new Dictionary<string, string> { { "EN", "3. Load ARES.mvba manually in MicroStation" }, { "FR", "3. Charger ARES.mvba manuellement dans MicroStation" } } },
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