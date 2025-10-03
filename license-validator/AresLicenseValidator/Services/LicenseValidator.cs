using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Security.Cryptography;
using Newtonsoft.Json;
using AresLicenseValidator.Models;

namespace AresLicenseValidator.Services
{
    internal class LicenseValidatorService
    {
        private const string LICENSE_FOLDER = "ARES_Licenses";
        private const string LICENSE_FILENAME = "ares_license.json";
        private const string LICENSE_VERSION = "1.0";

        // Clé publique RSA (sera remplacée par la vraie clé générée)
        private const string PUBLIC_KEY = @"<RSAKeyValue>
            <Modulus>VOTRE_CLE_PUBLIQUE_SERA_ICI_APRES_GENERATION</Modulus>
            <Exponent>AQAB</Exponent>
        </RSAKeyValue>";

        public string LastError { get; private set; } = "";

        public bool ValidateLicense()
        {
            try
            {
                LastError = "";

                // 1. Rechercher le fichier de licence
                var licensePath = FindLicenseFile();
                if (string.IsNullOrEmpty(licensePath))
                {
                    LastError = "License file not found on network drives";
                    return false;
                }

                // 2. Charger et parser le fichier
                var licenseData = LoadLicenseFile(licensePath);
                if (licenseData == null)
                {
                    LastError = "Invalid license file format";
                    return false;
                }

                // 3. Valider la signature cryptographique
                if (!ValidateSignature(licenseData))
                {
                    LastError = "Invalid license signature";
                    return false;
                }

                // 4. Valider l'environnement (domaine + utilisateur)
                if (!ValidateEnvironment(licenseData))
                {
                    return false; // LastError déjà défini dans ValidateEnvironment
                }

                return true;
            }
            catch (Exception ex)
            {
                LastError = $"License validation error: {ex.Message}";
                return false;
            }
        }

        public string GetLicenseInfo(out LicenseData licenseData)
        {
            licenseData = null;
            try
            {
                var licensePath = FindLicenseFile();
                if (string.IsNullOrEmpty(licensePath))
                    return "No license found on network";

                licenseData = LoadLicenseFile(licensePath);
                if (licenseData == null)
                    return "Invalid license file format";

                return $"Company: {licenseData.Company}\n" +
                       $"Domain: {licenseData.Domain}\n" +
                       $"Licensed Users: {licenseData.MaxUsers}\n" +
                       $"Installed: {licenseData.InstallationDate}\n" +
                       $"Admin: {licenseData.InstalledBy}";
            }
            catch (Exception ex)
            {
                return $"Error reading license: {ex.Message}";
            }
        }

        public string GetCurrentUser()
        {
            try
            {
                var domain = Environment.UserDomainName;
                var user = Environment.UserName;
                return $"{domain}\\{user}";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        private string FindLicenseFile()
        {
            // Chercher sur les lecteurs réseau mappés
            var networkDrives = new[] { "Z:", "Y:", "X:", "W:", "V:", "U:", "T:", "S:", "R:", "Q:", "P:", "O:", "N:", "M:", "L:", "K:" };

            foreach (var drive in networkDrives)
            {
                try
                {
                    var testPath = Path.Combine(drive + "\\", LICENSE_FOLDER, LICENSE_FILENAME);
                    if (File.Exists(testPath))
                        return testPath;
                }
                catch
                {
                    // Lecteur non accessible, continuer
                }
            }

            // Chercher sur des chemins UNC courants
            var uncPaths = new[]
            {
                @"\\server\shared\",
                @"\\fileserver\public\",
                @"\\nas\common\",
                @"\\srv01\data\",
                @"\\dc01\netlogon\"
            };

            foreach (var uncPath in uncPaths)
            {
                try
                {
                    var testPath = Path.Combine(uncPath, LICENSE_FOLDER, LICENSE_FILENAME);
                    if (File.Exists(testPath))
                        return testPath;
                }
                catch
                {
                    // Chemin UNC non accessible, continuer
                }
            }

            return null;
        }

        private LicenseData LoadLicenseFile(string path)
        {
            try
            {
                var jsonContent = File.ReadAllText(path, Encoding.UTF8);
                return JsonConvert.DeserializeObject<LicenseData>(jsonContent);
            }
            catch (Exception ex)
            {
                LastError = $"Error loading license file: {ex.Message}";
                return null;
            }
        }

        private bool ValidateSignature(LicenseData license)
        {
            try
            {
                using (var rsa = new RSACryptoServiceProvider())
                {
                    rsa.FromXmlString(PUBLIC_KEY);

                    // Reconstituer les données qui ont été signées
                    var dataToVerify = JsonConvert.SerializeObject(new
                    {
                        company = license.Company,
                        domain = license.Domain,
                        installed_by = license.InstalledBy,
                        installation_date = license.InstallationDate,
                        license_key = license.LicenseKey,
                        environment_hash = license.EnvironmentHash,
                        authorized_users = license.AuthorizedUsers,
                        max_users = license.MaxUsers
                    }, Formatting.None);

                    var dataBytes = Encoding.UTF8.GetBytes(dataToVerify);
                    var signatureBytes = Convert.FromBase64String(license.Signature);

                    return rsa.VerifyData(dataBytes, "SHA256", signatureBytes);
                }
            }
            catch (Exception ex)
            {
                LastError = $"Signature validation error: {ex.Message}";
                return false;
            }
        }

        private bool ValidateEnvironment(LicenseData license)
        {
            try
            {
                // 1. Vérifier le domaine Windows
                var currentDomain = Environment.UserDomainName;
                if (!string.Equals(currentDomain, license.Domain, StringComparison.OrdinalIgnoreCase))
                {
                    LastError = $"Domain mismatch: expected '{license.Domain}', got '{currentDomain}'";
                    return false;
                }

                // 2. Vérifier que l'utilisateur actuel est autorisé
                if (!ValidateCurrentUser(license.AuthorizedUsers, currentDomain))
                {
                    return false; // LastError déjà défini dans ValidateCurrentUser
                }

                // 3. Vérifier l'empreinte d'environnement
                var expectedHash = license.EnvironmentHash;
                var actualHash = CalculateEnvironmentHash(license.Company, license.Domain);

                if (!string.Equals(expectedHash, actualHash, StringComparison.Ordinal))
                {
                    LastError = "Environment fingerprint mismatch";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                LastError = $"Environment validation error: {ex.Message}";
                return false;
            }
        }

        private bool ValidateCurrentUser(string[] authorizedUsers, string currentDomain)
        {
            try
            {
                var currentUser = Environment.UserName;
                var fullCurrentUser = $"{currentDomain}\\{currentUser}".ToLower();

                // Vérifier si l'utilisateur est dans la liste
                var isAuthorized = authorizedUsers?.Any(user =>
                    string.Equals(user?.Trim(), fullCurrentUser, StringComparison.OrdinalIgnoreCase)) ?? false;

                if (!isAuthorized)
                {
                    LastError = $"User '{fullCurrentUser}' not authorized. Contact your administrator.";
                    LogUserAccess(fullCurrentUser, false);
                    return false;
                }

                LogUserAccess(fullCurrentUser, true);
                return true;
            }
            catch (Exception ex)
            {
                LastError = $"User validation error: {ex.Message}";
                return false;
            }
        }

        private string CalculateEnvironmentHash(string company, string domain)
        {
            var environmentData = $"{company}|{domain}|ARES_LICENSE_v1";

            using (var sha256 = SHA256.Create())
            {
                var hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(environmentData));
                var base64Hash = Convert.ToBase64String(hashBytes);
                return base64Hash.Substring(0, Math.Min(16, base64Hash.Length));
            }
        }

        private void LogUserAccess(string user, bool authorized)
        {
            try
            {
                var logPath = Path.Combine(Path.GetTempPath(), "ares_access.log");
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] User: {user}, Authorized: {authorized}";
                File.AppendAllText(logPath, logEntry + Environment.NewLine);
            }
            catch
            {
                // Silent fail pour les logs
            }
        }
    }
}