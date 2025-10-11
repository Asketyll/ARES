using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Security;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;

namespace AresInstaller
{
    public partial class AresInstallerForm : Form
    {
        #region Constants
        private const string GITHUB_RELEASES_URL = "https://api.github.com/repos/Asketyll/ARES/releases/latest";
        private const string INSTALL_PATH = @"C:\ARES\";
        private const string DLL_PATH = @"C:\ARES\Rsc\";
        private const string TEMP_DOWNLOAD_FOLDER = "ARES_Download";
        private const string TEMP_EXTRACT_FOLDER = "ARES_Extract";
        private const int DOTNET_FRAMEWORK_MIN_RELEASE = 461808;
        private const long MAX_FILE_SIZE = 100 * 1024 * 1024; // 100 MB per file
        private const long MAX_TOTAL_DOWNLOAD_SIZE = 500 * 1024 * 1024; // 500 MB total
        private static readonly string[] ALLOWED_EXTENSIONS = { ".dll", ".mvba", ".zip", ".tlb", ".sha256", ".md" };
        private const int MAX_RETRY_ATTEMPTS = 3;
        private const int BASE_RETRY_DELAY_MS = 1000;
        #endregion

        private string currentLanguage = "EN";
        private bool installationCompleted = false;
        private long totalDownloadedBytes = 0;

        #region Security Helper Classes
        private class FileIntegrityValidator
        {
            private Dictionary<string, string> expectedHashes;

            public FileIntegrityValidator()
            {
                expectedHashes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            public void LoadHashesFromFile(string hashFilePath)
            {
                if (!File.Exists(hashFilePath))
                    return;

                foreach (var line in File.ReadAllLines(hashFilePath))
                {
                    var parts = line.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length >= 2)
                    {
                        expectedHashes[parts[1]] = parts[0].ToLowerInvariant();
                    }
                }
            }

            public bool VerifyFile(string filePath)
            {
                var fileName = Path.GetFileName(filePath);
                if (!expectedHashes.ContainsKey(fileName))
                    return false;

                using (var sha256 = SHA256.Create())
                using (var stream = File.OpenRead(filePath))
                {
                    var hash = BitConverter.ToString(sha256.ComputeHash(stream))
                        .Replace("-", "").ToLowerInvariant();
                    return hash.Equals(expectedHashes[fileName], StringComparison.OrdinalIgnoreCase);
                }
            }

            public bool HasHash(string fileName)
            {
                return expectedHashes.ContainsKey(fileName);
            }
        }

        private class PathValidator
        {
            public static bool IsValidFileName(string fileName)
            {
                if (string.IsNullOrWhiteSpace(fileName))
                    return false;

                if (fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                    return false;

                if (fileName.Contains("..") || fileName.Contains("/") || fileName.Contains("\\"))
                    return false;

                var extension = Path.GetExtension(fileName).ToLowerInvariant();
                return ALLOWED_EXTENSIONS.Contains(extension);
            }

            public static string SanitizeLogMessage(string message)
            {
                if (string.IsNullOrEmpty(message))
                    return string.Empty;

                return message
                    .Replace("\r", "")
                    .Replace("\n", " ")
                    .Replace("\0", "")
                    .Substring(0, Math.Min(message.Length, 500));
            }
        }
        #endregion

        #region Version Management
        private string ExtractVersionFromFilename(string filename)
        {
            try
            {
                var nameWithoutExt = Path.GetFileNameWithoutExtension(filename);
                var parts = nameWithoutExt.Split('-');

                if (parts.Length >= 2)
                {
                    return parts[parts.Length - 1];
                }
            }
            catch
            {
                // Ignore parsing errors
            }

            return string.Empty;
        }

        private bool IsSameVersion(string sourceFile, string targetFile)
        {
            if (!File.Exists(targetFile))
            {
                return false;
            }

            var sourceVersion = ExtractVersionFromFilename(sourceFile);
            var targetVersion = ExtractVersionFromFilename(targetFile);

            if (string.IsNullOrEmpty(sourceVersion) || string.IsNullOrEmpty(targetVersion))
            {
                return false;
            }

            return sourceVersion.Equals(targetVersion, StringComparison.OrdinalIgnoreCase);
        }

        private string FindExistingDllWithBaseName(string dllBaseName)
        {
            try
            {
                var searchPattern = $"{dllBaseName}-*.dll";
                var existingFiles = Directory.GetFiles(DLL_PATH, searchPattern);

                if (existingFiles.Length > 0)
                {
                    return existingFiles[0];
                }
            }
            catch
            {
                // Ignore errors
            }

            return null;
        }
        #endregion

        #region UI Controls
        private ProgressBar progressBar;
        private Label statusLabel;
        private Button installButton;
        private RichTextBox logTextBox;
        #endregion

        #region Constructor
        public AresInstallerForm(string language)
        {
            currentLanguage = language;
            InitializeComponent();
            SetupCustomControls();
            ApplyTranslations();
        }
        #endregion

        #region UI Setup
        private void SetupCustomControls()
        {
            this.Size = new System.Drawing.Size(600, 500);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(540, 23),
                Style = ProgressBarStyle.Continuous
            };
            this.Controls.Add(progressBar);

            statusLabel = new Label
            {
                Location = new System.Drawing.Point(20, 50),
                Size = new System.Drawing.Size(540, 20)
            };
            this.Controls.Add(statusLabel);

            installButton = new Button
            {
                Location = new System.Drawing.Point(250, 80),
                Size = new System.Drawing.Size(100, 30)
            };
            installButton.Click += InstallButton_Click;
            this.Controls.Add(installButton);

            logTextBox = new RichTextBox
            {
                Location = new System.Drawing.Point(20, 120),
                Size = new System.Drawing.Size(540, 320),
                ReadOnly = true,
                Font = new System.Drawing.Font("Consolas", 9)
            };
            this.Controls.Add(logTextBox);
        }
        #endregion

        #region Event Handlers
        private async void InstallButton_Click(object sender, EventArgs e)
        {
            if (installationCompleted)
            {
                this.Close();
                return;
            }

            installButton.Enabled = false;

            try
            {
                await PerformInstallation();
                MessageBox.Show(
                    Translations.Get("InstallSuccess", currentLanguage),
                    Translations.Get("InstallationComplete", currentLanguage),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );

                installationCompleted = true;
                installButton.Text = Translations.Get("ExitButton", currentLanguage);
                installButton.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    Translations.Format("InstallError", currentLanguage, ex.Message),
                    Translations.Get("InstallationError", currentLanguage),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                LogMessage($"ERROR: {ex.GetType().Name} - {ex.Message}");

                installationCompleted = true;
                installButton.Text = Translations.Get("ExitButton", currentLanguage);
                installButton.Enabled = true;
            }
        }
        #endregion

        #region Main Installation Logic
        private async Task PerformInstallation()
        {
            const int totalSteps = 7;
            progressBar.Value = 0;
            progressBar.Maximum = totalSteps;

            try
            {
                UpdateStatus("CheckingPrerequisites");
                await CheckPrerequisites();
                IncrementProgress();

                UpdateStatus("CreatingDirectories");
                CreateDirectories();
                IncrementProgress();

                UpdateStatus("Downloading");
                await DownloadFromGitHub();
                IncrementProgress();

                UpdateStatus("Extracting");
                await ExtractDownloadedFiles();
                IncrementProgress();

                UpdateStatus("InstallingProject");
                await CopyMVBAProject();
                IncrementProgress();

                UpdateStatus("RegisteringCOM");
                await RegisterDLLs();
                IncrementProgress();

                UpdateStatus("InstallationCompleted");
                await CleanupTemporaryFiles();
                LogInstallationSummary();
                IncrementProgress();
            }
            catch (Exception ex)
            {
                UpdateStatus("InstallationError");
                LogMessage($"ERROR: {ex.Message}");
                await CleanupTemporaryFiles();
                throw;
            }
        }

        private void IncrementProgress()
        {
            progressBar.Value++;
        }

        private void LogInstallationSummary()
        {
            LogMessage(Translations.Get("InstallationSummary", currentLanguage));
            LogMessage(Translations.Format("MainProject", currentLanguage, INSTALL_PATH));
            LogMessage(Translations.Format("DLLComponents", currentLanguage, DLL_PATH));
            LogMessage(Translations.Get("COMRegistered", currentLanguage));
            LogMessage("");
            LogMessage(Translations.Get("NextSteps", currentLanguage));
            LogMessage(Translations.Get("Step1", currentLanguage));
            LogMessage(Translations.Get("Step2", currentLanguage));
            LogMessage(Translations.Get("Step3", currentLanguage));
        }
        #endregion

        #region Prerequisites Check
        private async Task CheckPrerequisites()
        {
            LogMessage(Translations.Get("PrerequisitesCheck", currentLanguage));

            await Task.Delay(500);

            if (!IsRunningAsAdministrator())
            {
                LogMessage(Translations.Get("NotRunningAsAdmin", currentLanguage));
            }

            if (IsDotNetFrameworkInstalled())
            {
                LogMessage(Translations.Get("DotNetAvailable", currentLanguage));
            }
            else
            {
                throw new InvalidOperationException(Translations.Get("DotNetRequired", currentLanguage));
            }

            // Verify HTTPS is used
            if (!GITHUB_RELEASES_URL.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                throw new SecurityException("Only HTTPS connections are allowed for downloads");
            }

            LogMessage("");
        }

        private bool IsRunningAsAdministrator()
        {
            try
            {
                var identity = WindowsIdentity.GetCurrent();
                var principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch
            {
                return false;
            }
        }

        private bool IsDotNetFrameworkInstalled()
        {
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\"))
                {
                    if (key?.GetValue("Release") is int release)
                    {
                        return release >= DOTNET_FRAMEWORK_MIN_RELEASE;
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
        }
        #endregion

        #region Directory Management
        private void CreateDirectories()
        {
            LogMessage(Translations.Get("CreatingDirs", currentLanguage));

            try
            {
                var directoriesToCreate = new[]
                {
                    INSTALL_PATH,
                    DLL_PATH,
                    Path.Combine(INSTALL_PATH, "Backup"),
                };

                foreach (var directory in directoriesToCreate)
                {
                    Directory.CreateDirectory(directory);
                    LogMessage(Translations.Format("Created", currentLanguage, directory));
                }
            }
            catch (Exception ex)
            {
                throw new DirectoryCreationException(Translations.Format("FailedCreateDirs", currentLanguage, ex.Message), ex);
            }

            LogMessage("");
        }

        private string GetSecureTempPath(string folderName)
        {
            var uniqueFolderName = $"{folderName}_{Guid.NewGuid():N}";
            var tempPath = Path.Combine(Path.GetTempPath(), uniqueFolderName);

            var dirInfo = Directory.CreateDirectory(tempPath);

            try
            {
                var dirSecurity = dirInfo.GetAccessControl();
                dirSecurity.SetAccessRuleProtection(true, false);

                var currentUser = WindowsIdentity.GetCurrent();
                var rule = new FileSystemAccessRule(
                    currentUser.User,
                    FileSystemRights.FullControl,
                    InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
                    PropagationFlags.None,
                    AccessControlType.Allow);

                dirSecurity.AddAccessRule(rule);
                dirInfo.SetAccessControl(dirSecurity);
            }
            catch
            {
                // Continue even if ACL setup fails
            }

            return tempPath;
        }
        #endregion

        #region GitHub Download
        private async Task DownloadFromGitHub()
        {
            LogMessage(Translations.Get("DownloadingFromGitHub", currentLanguage));

            using (var client = new HttpClient())
            {
                ConfigureHttpClient(client);

                try
                {
                    LogMessage(Translations.Get("FetchingReleaseInfoAPI", currentLanguage));
                    var releaseInfo = await RetryWithExponentialBackoff(() => GetLatestReleaseInfo(client));

                    LogMessage(Translations.Get("ParsingReleaseAssets", currentLanguage));
                    await DownloadReleaseAssets(client, releaseInfo);
                }
                catch (Exception ex)
                {
                    LogMessage($"DOWNLOAD ERROR: {ex.GetType().Name}");
                    LogMessage($"Message: {ex.Message}");
                    throw new DownloadException(Translations.Format("FailedDownload", currentLanguage, ex.Message), ex);
                }
            }

            LogMessage("");
        }

        private static void ConfigureHttpClient(HttpClient client)
        {
            client.DefaultRequestHeaders.Add("User-Agent", "ARES-Installer");
            client.Timeout = TimeSpan.FromMinutes(10);
        }

        private async Task<string> GetLatestReleaseInfo(HttpClient client)
        {
            try
            {
                var releaseResponse = await client.GetStringAsync(GITHUB_RELEASES_URL);

                var release = JObject.Parse(releaseResponse);
                var tagName = release["tag_name"]?.ToString();

                if (string.IsNullOrEmpty(tagName))
                {
                    LogMessage("WARNING: Could not find tag_name in response");
                }
                else
                {
                    LogMessage(Translations.Format("LatestVersion", currentLanguage, tagName));
                }

                return releaseResponse;
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: {ex.Message}");
                throw;
            }
        }

        private async Task DownloadReleaseAssets(HttpClient client, string releaseResponse)
        {
            try
            {
                var release = JObject.Parse(releaseResponse);
                var assets = release["assets"] as JArray;

                if (assets == null || assets.Count == 0)
                {
                    throw new DownloadException(Translations.Get("NoAssetsFound", currentLanguage));
                }

                LogMessage(Translations.Format("FoundAssets", currentLanguage, assets.Count));

                var downloadPath = GetSecureTempPath(TEMP_DOWNLOAD_FOLDER);
                Directory.CreateDirectory(downloadPath);

                // Look for checksum file first
                var checksumAsset = assets.FirstOrDefault(a =>
                    a["name"]?.ToString().EndsWith(".sha256", StringComparison.OrdinalIgnoreCase) == true);

                FileIntegrityValidator validator = null;
                if (checksumAsset != null)
                {
                    var checksumFileName = checksumAsset["name"]?.ToString();
                    var checksumUrl = checksumAsset["browser_download_url"]?.ToString();

                    if (!string.IsNullOrEmpty(checksumFileName) && !string.IsNullOrEmpty(checksumUrl))
                    {
                        await DownloadFile(client, checksumFileName, checksumUrl, downloadPath);
                        validator = new FileIntegrityValidator();
                        validator.LoadHashesFromFile(Path.Combine(downloadPath, checksumFileName));
                        LogMessage("Checksum file loaded for integrity verification");
                    }
                }

                foreach (var asset in assets)
                {
                    var fileName = asset["name"]?.ToString();
                    var downloadUrl = asset["browser_download_url"]?.ToString();

                    if (string.IsNullOrEmpty(fileName) || string.IsNullOrEmpty(downloadUrl))
                        continue;

                    // Skip checksum file itself
                    if (fileName.EndsWith(".sha256", StringComparison.OrdinalIgnoreCase))
                        continue;

                    // Validate filename
                    if (!PathValidator.IsValidFileName(fileName))
                    {
                        LogMessage($"SECURITY: Skipped invalid filename: {fileName}");
                        continue;
                    }

                    await DownloadFile(client, fileName, downloadUrl, downloadPath);

                    // Verify hash if available
                    if (validator != null && validator.HasHash(fileName))
                    {
                        var filePath = Path.Combine(downloadPath, fileName);
                        if (!validator.VerifyFile(filePath))
                        {
                            File.Delete(filePath);
                            throw new SecurityException($"Hash verification failed for {fileName}");
                        }
                        LogMessage($"  ✓ Hash verified for {fileName}");
                    }
                }

                var downloadedFiles = Directory.GetFiles(downloadPath);
                LogMessage(Translations.Format("DownloadComplete", currentLanguage, downloadedFiles.Length));
            }
            catch (Exception ex)
            {
                throw new DownloadException(Translations.Format("FailedDownloadAssets", currentLanguage, ex.Message), ex);
            }
        }

        private async Task DownloadFile(HttpClient client, string fileName, string downloadUrl, string downloadPath)
        {
            LogMessage(Translations.Format("DownloadingFile", currentLanguage, fileName));

            try
            {
                using (var response = await client.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead))
                {
                    response.EnsureSuccessStatusCode();

                    if (response.Content.Headers.ContentLength.HasValue)
                    {
                        var fileSize = response.Content.Headers.ContentLength.Value;

                        if (fileSize > MAX_FILE_SIZE)
                        {
                            throw new SecurityException($"File {fileName} exceeds maximum size ({MAX_FILE_SIZE} bytes)");
                        }

                        if (totalDownloadedBytes + fileSize > MAX_TOTAL_DOWNLOAD_SIZE)
                        {
                            throw new SecurityException($"Total download size exceeds limit ({MAX_TOTAL_DOWNLOAD_SIZE} bytes)");
                        }
                    }

                    var fileBytes = await response.Content.ReadAsByteArrayAsync();

                    if (response.Content.Headers.ContentLength.HasValue &&
                        fileBytes.Length != response.Content.Headers.ContentLength.Value)
                    {
                        throw new SecurityException($"File size mismatch for {fileName}");
                    }

                    totalDownloadedBytes += fileBytes.Length;

                    var filePath = Path.Combine(downloadPath, fileName);
                    File.WriteAllBytes(filePath, fileBytes);

                    if (!File.Exists(filePath))
                    {
                        throw new IOException($"File was not created: {filePath}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: {ex.GetType().Name} - {ex.Message}");
                throw;
            }
        }

        private async Task<T> RetryWithExponentialBackoff<T>(Func<Task<T>> operation)
        {
            for (int i = 0; i < MAX_RETRY_ATTEMPTS; i++)
            {
                try
                {
                    return await operation();
                }
                catch (Exception ex) when (i < MAX_RETRY_ATTEMPTS - 1)
                {
                    var delay = BASE_RETRY_DELAY_MS * (int)Math.Pow(2, i);
                    LogMessage($"Retry {i + 1}/{MAX_RETRY_ATTEMPTS} after {delay}ms: {ex.Message}");
                    await Task.Delay(delay);
                }
            }

            return await operation();
        }
        #endregion

        #region File Extraction
        private async Task ExtractDownloadedFiles()
        {
            LogMessage(Translations.Get("ExtractingFiles", currentLanguage));

            try
            {
                var tempBase = Path.GetTempPath();
                var downloadFolders = Directory.GetDirectories(tempBase, $"{TEMP_DOWNLOAD_FOLDER}_*");

                if (downloadFolders.Length == 0)
                {
                    throw new ExtractionException("Download folder not found");
                }

                var downloadPath = downloadFolders.OrderByDescending(d => Directory.GetCreationTime(d)).First();
                var extractPath = GetSecureTempPath(TEMP_EXTRACT_FOLDER);

                Directory.CreateDirectory(extractPath);

                await ExtractZipFiles(downloadPath, extractPath);
                CopyNonZipFiles(downloadPath, extractPath);
            }
            catch (Exception ex)
            {
                throw new ExtractionException(Translations.Format("FailedExtract", currentLanguage, ex.Message), ex);
            }

            LogMessage("");
        }

        private async Task ExtractZipFiles(string downloadPath, string extractPath)
        {
            if (!Directory.Exists(downloadPath))
                return;

            var zipFiles = Directory.GetFiles(downloadPath, "*.zip");

            foreach (var zipFile in zipFiles)
            {
                LogMessage(Translations.Format("ExtractingZip", currentLanguage, Path.GetFileName(zipFile)));

                using (var archive = ZipFile.OpenRead(zipFile))
                {
                    foreach (var entry in archive.Entries)
                    {
                        if (!string.IsNullOrEmpty(entry.Name))
                        {
                            await ExtractZipEntry(entry, extractPath);
                        }
                    }
                }
            }
        }

        private async Task ExtractZipEntry(ZipArchiveEntry entry, string extractPath)
        {
            // SECURITY FIX: Validate against path traversal (Zip Slip)
            // Resolve the full path where the entry would be extracted
            string fullEntryPath = Path.GetFullPath(Path.Combine(extractPath, entry.FullName));

            // Ensure extraction directory path ends with separator for accurate comparison
            string fullExtractPath = Path.GetFullPath(extractPath);
            if (!fullExtractPath.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                fullExtractPath += Path.DirectorySeparatorChar;
            }

            // Check if the resolved path is within the extraction directory
            if (!fullEntryPath.StartsWith(fullExtractPath, StringComparison.OrdinalIgnoreCase))
            {
                LogMessage($"SECURITY: Blocked path traversal attempt: {entry.FullName}");
                throw new SecurityException($"Invalid archive entry path detected: {entry.FullName}");
            }

            // Validate filename
            var fileName = Path.GetFileName(entry.Name);
            if (!PathValidator.IsValidFileName(fileName))
            {
                LogMessage($"SECURITY: Skipped invalid filename in archive: {fileName}");
                return;
            }

            // CRITICAL: Use fullEntryPath (sanitized) for all file operations
            var directory = Path.GetDirectoryName(fullEntryPath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            entry.ExtractToFile(fullEntryPath, true);
            LogMessage(Translations.Format("Extracted", currentLanguage, entry.Name));

            await Task.Delay(50);
        }

        private void CopyNonZipFiles(string downloadPath, string extractPath)
        {
            if (!Directory.Exists(downloadPath))
                return;

            var otherFiles = Directory.GetFiles(downloadPath)
                .Where(f => !f.EndsWith(".zip", StringComparison.OrdinalIgnoreCase));

            foreach (var file in otherFiles)
            {
                var fileName = Path.GetFileName(file);

                if (!PathValidator.IsValidFileName(fileName))
                {
                    LogMessage($"SECURITY: Skipped invalid filename: {fileName}");
                    continue;
                }

                var destPath = Path.Combine(extractPath, fileName);
                File.Copy(file, destPath, true);
                LogMessage(Translations.Format("Copied", currentLanguage, fileName));
            }
        }
        #endregion

        #region DLL Registration
        private async Task RegisterDLLs()
        {
            LogMessage(Translations.Get("RegisteringComponents", currentLanguage));

            try
            {
                var tempBase = Path.GetTempPath();
                var extractFolders = Directory.GetDirectories(tempBase, $"{TEMP_EXTRACT_FOLDER}_*");

                if (extractFolders.Length == 0)
                {
                    throw new RegistrationException("Extract folder not found");
                }

                var extractPath = extractFolders.OrderByDescending(d => Directory.GetCreationTime(d)).First();

                await CopyDLLsToInstallPath(extractPath);

                LogMessage(Translations.Get("SearchingValidator", currentLanguage));
                var validatorDll = FindAresLicenseValidatorDll();

                if (string.IsNullOrEmpty(validatorDll))
                {
                    throw new FileNotFoundException(Translations.Get("ValidatorNotFound", currentLanguage));
                }

                if (!File.Exists(validatorDll))
                {
                    throw new FileNotFoundException(Translations.Format("ValidatorNotFoundAtPath", currentLanguage, validatorDll));
                }

                LogMessage(Translations.Format("FoundDLL", currentLanguage, validatorDll));
                await RegisterSingleDLL(validatorDll);
                LogMessage(Translations.Get("COMRegistrationComplete", currentLanguage));
            }
            catch (Exception ex)
            {
                throw new RegistrationException(Translations.Format("FailedRegisterDLLs", currentLanguage, ex.Message), ex);
            }

            LogMessage("");
        }

        private string FindAresLicenseValidatorDll()
        {
            try
            {
                if (!Directory.Exists(DLL_PATH))
                {
                    return null;
                }

                var searchPatterns = new[]
                {
                    "AresLicenseValidator.dll",
                    "AresLicenseValidator-*.dll"
                };

                foreach (var pattern in searchPatterns)
                {
                    var files = Directory.GetFiles(DLL_PATH, pattern);

                    if (files.Length > 0)
                    {
                        return files[0];
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR in FindAresLicenseValidatorDll: {ex.Message}");
            }

            return null;
        }

        private async Task CopyDLLsToInstallPath(string sourcePath)
        {
            var dllFiles = Directory.GetFiles(sourcePath, "*.dll");

            foreach (var sourceDll in dllFiles)
            {
                var fileName = Path.GetFileName(sourceDll);

                if (!PathValidator.IsValidFileName(fileName))
                {
                    LogMessage($"SECURITY: Skipped invalid DLL filename: {fileName}");
                    continue;
                }

                await CopySingleDLL(sourcePath, fileName);
            }
        }

        private async Task CopySingleDLL(string sourcePath, string dllFileName)
        {
            var sourceDll = Path.Combine(sourcePath, dllFileName);

            LogMessage(Translations.Format("ProcessingDLL", currentLanguage, dllFileName));

            if (!File.Exists(sourceDll))
            {
                return;
            }

            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(dllFileName);
            var dllBaseName = fileNameWithoutExt;

            var lastDashIndex = fileNameWithoutExt.LastIndexOf('-');
            if (lastDashIndex > 0)
            {
                dllBaseName = fileNameWithoutExt.Substring(0, lastDashIndex);
            }

            var existingDll = FindExistingDllWithBaseName(dllBaseName);

            if (!string.IsNullOrEmpty(existingDll))
            {
                if (lastDashIndex > 0 && IsSameVersion(sourceDll, existingDll))
                {
                    LogMessage(Translations.Get("SameVersionInstalled", currentLanguage));
                    return;
                }

                var backupPath = Path.Combine(INSTALL_PATH, "Backup",
                    $"{Path.GetFileName(existingDll)}.backup_{DateTime.Now:yyyyMMdd_HHmmss}");

                File.Move(existingDll, backupPath);
                LogMessage(Translations.Format("BackedUpOldVersion", currentLanguage, Path.GetFileName(backupPath)));

                var existingTlb = Path.ChangeExtension(existingDll, ".tlb");
                if (File.Exists(existingTlb))
                {
                    var tlbBackupPath = Path.Combine(INSTALL_PATH, "Backup",
                        $"{Path.GetFileName(existingTlb)}.backup_{DateTime.Now:yyyyMMdd_HHmmss}");
                    File.Move(existingTlb, tlbBackupPath);
                    LogMessage(Translations.Format("BackedUpOldTLB", currentLanguage, Path.GetFileName(tlbBackupPath)));
                }
            }

            var targetDll = Path.Combine(DLL_PATH, dllFileName);
            File.Copy(sourceDll, targetDll, true);

            if (File.Exists(targetDll))
            {
                LogMessage(Translations.Format("CopiedDLL", currentLanguage, dllFileName, new FileInfo(targetDll).Length));
            }
            else
            {
                throw new IOException(Translations.Format("FailedCopyDLL", currentLanguage, targetDll));
            }

            var sourceTlb = Path.ChangeExtension(sourceDll, ".tlb");
            if (File.Exists(sourceTlb))
            {
                var targetTlb = Path.Combine(DLL_PATH, Path.GetFileName(sourceTlb));
                File.Copy(sourceTlb, targetTlb, true);

                if (File.Exists(targetTlb))
                {
                    LogMessage(Translations.Format("CopiedTLB", currentLanguage, Path.GetFileName(sourceTlb)));
                }
            }

            await Task.Delay(100);
        }

        private async Task RegisterSingleDLL(string dllPath)
        {
            LogMessage(Translations.Format("RegisteringDLL", currentLanguage, Path.GetFileName(dllPath)));

            // SECURITY: Validate DLL path is within expected directory
            string fullDllPath = Path.GetFullPath(dllPath);
            string fullDllDirectory = Path.GetFullPath(DLL_PATH);

            if (!fullDllDirectory.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                fullDllDirectory += Path.DirectorySeparatorChar;
            }

            if (!fullDllPath.StartsWith(fullDllDirectory, StringComparison.OrdinalIgnoreCase))
            {
                throw new SecurityException("DLL path outside expected directory");
            }

            if (!File.Exists(dllPath))
            {
                throw new FileNotFoundException($"DLL not found: {dllPath}");
            }

            var regasmPath = FindRegAsmPath();
            if (string.IsNullOrEmpty(regasmPath))
            {
                throw new FileNotFoundException(Translations.Get("RegAsmNotFound", currentLanguage));
            }

            await ExecuteRegAsm(regasmPath, dllPath);
        }

        private async Task ExecuteRegAsm(string regasmPath, string dllPath)
        {
            using (var process = new Process())
            {
                ConfigureRegAsmProcess(process, regasmPath, dllPath);

                try
                {
                    process.Start();

                    var output = await process.StandardOutput.ReadToEndAsync();
                    var error = await process.StandardError.ReadToEndAsync();

                    process.WaitForExit();

                    HandleRegAsmResult(process.ExitCode, output, error);
                }
                catch (Exception ex)
                {
                    throw new RegistrationException(Translations.Format("FailedExecuteRegAsm", currentLanguage, ex.Message), ex);
                }
            }
        }

        private static void ConfigureRegAsmProcess(Process process, string regasmPath, string dllPath)
        {
            process.StartInfo = new ProcessStartInfo
            {
                FileName = regasmPath,
                Arguments = $"\"{dllPath}\" /tlb /codebase",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
        }

        private void HandleRegAsmResult(int exitCode, string output, string error)
        {
            if (exitCode == 0)
            {
                LogMessage(Translations.Get("DLLRegisteredSuccess", currentLanguage));
            }
            else
            {
                throw new RegistrationException(Translations.Format("RegAsmFailed", currentLanguage, exitCode, error));
            }
        }

        private string FindRegAsmPath()
        {
            var possiblePaths = new[]
            {
                @"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe",
                @"C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm.exe"
            };

            return possiblePaths.FirstOrDefault(File.Exists);
        }
        #endregion

        #region Project Installation
        private async Task CopyMVBAProject()
        {
            LogMessage(Translations.Get("InstallingARES", currentLanguage));

            try
            {
                var tempBase = Path.GetTempPath();
                var extractFolders = Directory.GetDirectories(tempBase, $"{TEMP_EXTRACT_FOLDER}_*");

                if (extractFolders.Length == 0)
                {
                    throw new InstallationException("Extract folder not found");
                }

                var extractPath = extractFolders.OrderByDescending(d => Directory.GetCreationTime(d)).First();

                await CopyMVBAFile(extractPath);
                await Task.Delay(500);
            }
            catch (Exception ex)
            {
                throw new InstallationException(Translations.Format("FailedCopyProject", currentLanguage, ex.Message), ex);
            }

            LogMessage("");
        }

        private async Task CopyMVBAFile(string extractPath)
        {
            var mvbaSource = Path.Combine(extractPath, "ARES.mvba");

            if (File.Exists(mvbaSource))
            {
                var mvbaTarget = Path.Combine(INSTALL_PATH, "ARES.mvba");
                File.Copy(mvbaSource, mvbaTarget, true);
                LogMessage(Translations.Format("CopiedMVBA", currentLanguage, mvbaTarget));
            }
            else
            {
                LogMessage(Translations.Get("MVBANotFound", currentLanguage));
            }

            await Task.Delay(100);
        }
        #endregion

        #region Cleanup
        private async Task CleanupTemporaryFiles()
        {
            LogMessage(Translations.Get("CleaningUp", currentLanguage));

            try
            {
                var tempBase = Path.GetTempPath();
                var foldersToClean = new[]
                {
                    $"{TEMP_DOWNLOAD_FOLDER}_*",
                    $"{TEMP_EXTRACT_FOLDER}_*"
                };

                foreach (var pattern in foldersToClean)
                {
                    var matchingFolders = Directory.GetDirectories(tempBase, pattern);
                    foreach (var folder in matchingFolders)
                    {
                        try
                        {
                            Directory.Delete(folder, true);
                            await Task.Delay(100);
                        }
                        catch
                        {
                            // Continue cleanup even if some folders fail
                        }
                    }
                }

                LogMessage(Translations.Get("CleanupCompleted", currentLanguage));
            }
            catch (Exception ex)
            {
                LogMessage(Translations.Format("CleanupWarning", currentLanguage, ex.Message));
            }

            LogMessage("");
        }
        #endregion

        #region Utility Methods
        private string GetTempPath(string folderName)
        {
            return GetSecureTempPath(folderName);
        }

        private void UpdateStatus(string key)
        {
            var message = Translations.Get(key, currentLanguage);
            if (statusLabel.InvokeRequired)
            {
                statusLabel.Invoke(new Action(() => statusLabel.Text = message));
            }
            else
            {
                statusLabel.Text = message;
            }
            Application.DoEvents();
        }

        private void LogMessage(string message)
        {
            var sanitized = PathValidator.SanitizeLogMessage(message);
            var logEntry = $"{DateTime.Now:HH:mm:ss} - {sanitized}\n";

            if (logTextBox.InvokeRequired)
            {
                logTextBox.Invoke(new Action(() => {
                    logTextBox.AppendText(logEntry);
                    logTextBox.ScrollToCaret();
                }));
            }
            else
            {
                logTextBox.AppendText(logEntry);
                logTextBox.ScrollToCaret();
            }
            Application.DoEvents();
        }

        private void ApplyTranslations()
        {
            this.Text = Translations.Get("WindowTitle", currentLanguage);
            statusLabel.Text = Translations.Get("ReadyToInstall", currentLanguage);
            installButton.Text = Translations.Get("InstallButton", currentLanguage);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!installationCompleted && installButton.Enabled == false)
            {
                MessageBox.Show(
                    Translations.Get("InstallationInProgress", currentLanguage),
                    this.Text,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                e.Cancel = true;
            }

            base.OnFormClosing(e);
        }
        #endregion
    }

    #region Custom Exceptions
    public class DirectoryCreationException : Exception
    {
        public DirectoryCreationException(string message) : base(message) { }
        public DirectoryCreationException(string message, Exception innerException) : base(message, innerException) { }
    }

    public class DownloadException : Exception
    {
        public DownloadException(string message) : base(message) { }
        public DownloadException(string message, Exception innerException) : base(message, innerException) { }
    }

    public class ExtractionException : Exception
    {
        public ExtractionException(string message) : base(message) { }
        public ExtractionException(string message, Exception innerException) : base(message, innerException) { }
    }

    public class RegistrationException : Exception
    {
        public RegistrationException(string message) : base(message) { }
        public RegistrationException(string message, Exception innerException) : base(message, innerException) { }
    }

    public class InstallationException : Exception
    {
        public InstallationException(string message) : base(message) { }
        public InstallationException(string message, Exception innerException) : base(message, innerException) { }
    }
    #endregion
}