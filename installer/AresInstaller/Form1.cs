using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
        private const int DOTNET_FRAMEWORK_MIN_RELEASE = 461808; // .NET Framework 4.7.2
        #endregion

        private string currentLanguage = "EN";
        private bool installationCompleted = false;

        #region Version Management
        private string ExtractVersionFromFilename(string filename)
        {
            // Extract version from format: AresLicenseValidator-1.0.0.dll
            try
            {
                var nameWithoutExt = Path.GetFileNameWithoutExtension(filename);
                var parts = nameWithoutExt.Split('-');

                if (parts.Length >= 2)
                {
                    return parts[parts.Length - 1]; // Return last part (version)
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
            // Find any DLL matching the base name pattern (e.g., AresLicenseValidator-*.dll)
            try
            {
                var searchPattern = $"{dllBaseName}-*.dll";
                var existingFiles = Directory.GetFiles(DLL_PATH, searchPattern);

                if (existingFiles.Length > 0)
                {
                    return existingFiles[0]; // Return first match
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
            // Window configuration
            this.Size = new System.Drawing.Size(600, 500);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            // ProgressBar
            progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(540, 23),
                Style = ProgressBarStyle.Continuous
            };
            this.Controls.Add(progressBar);

            // Status Label
            statusLabel = new Label
            {
                Location = new System.Drawing.Point(20, 50),
                Size = new System.Drawing.Size(540, 20)
            };
            this.Controls.Add(statusLabel);

            // Install/Close Button
            installButton = new Button
            {
                Location = new System.Drawing.Point(250, 80),
                Size = new System.Drawing.Size(100, 30)
            };
            installButton.Click += InstallButton_Click;
            this.Controls.Add(installButton);

            // Log TextBox
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
                // Close/Exit button behavior
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

                // Change button to "Close" on success
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
                LogMessage($"ERROR: {ex}");

                // Change button to "Exit" on error
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
            LogMessage("=== Installation Summary ===");
            LogMessage($"Main project: {INSTALL_PATH}");
            LogMessage($"DLL components: {DLL_PATH}");
            LogMessage("COM components registered");
            LogMessage("");
            LogMessage("Next steps:");
            LogMessage("Load ARES.mvba manually in MicroStation");
        }
        #endregion

        #region Prerequisites Check
        private async Task CheckPrerequisites()
        {
            LogMessage("=== Prerequisites Check ===");

            await Task.Delay(500);

            // Check administrator privileges
            if (!IsRunningAsAdministrator())
            {
                LogMessage("WARNING: Not running as Administrator - some features may not work");
            }
            else
            {
                LogMessage("Running as Administrator");
            }

            // Check .NET Framework
            if (IsDotNetFrameworkInstalled())
            {
                LogMessage(".NET Framework 4.7.2+ available");
            }
            else
            {
                throw new InvalidOperationException(".NET Framework 4.7.2 or higher is required");
            }

            LogMessage("");
        }

        private bool IsRunningAsAdministrator()
        {
            try
            {
                var identity = System.Security.Principal.WindowsIdentity.GetCurrent();
                var principal = new System.Security.Principal.WindowsPrincipal(identity);
                return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
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
            LogMessage("=== Creating Directories ===");

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
                    LogMessage($"Created: {directory}");
                }
            }
            catch (Exception ex)
            {
                throw new DirectoryCreationException($"Failed to create directories: {ex.Message}", ex);
            }

            LogMessage("");
        }
        #endregion

        #region GitHub Download
        private async Task DownloadFromGitHub()
        {
            LogMessage("=== Downloading from GitHub ===");

            using (var client = new HttpClient())
            {
                ConfigureHttpClient(client);

                try
                {
                    var releaseInfo = await GetLatestReleaseInfo(client);
                    await DownloadReleaseAssets(client, releaseInfo);
                }
                catch (Exception ex)
                {
                    throw new DownloadException($"Failed to download from GitHub: {ex.Message}", ex);
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
            LogMessage("Fetching release information...");
            var releaseResponse = await client.GetStringAsync(GITHUB_RELEASES_URL);

            var tagName = ExtractJsonValue(releaseResponse, "tag_name");
            LogMessage($"Latest version: {tagName}");

            return releaseResponse;
        }

        private async Task DownloadReleaseAssets(HttpClient client, string releaseResponse)
        {
            var assetsSection = ExtractJsonSection(releaseResponse, "assets");
            var downloadUrls = ExtractDownloadUrls(assetsSection);

            var downloadPath = GetTempPath(TEMP_DOWNLOAD_FOLDER);
            Directory.CreateDirectory(downloadPath);

            foreach (var (fileName, downloadUrl) in downloadUrls)
            {
                await DownloadFile(client, fileName, downloadUrl, downloadPath);
            }

            LogMessage($"Files saved to: {downloadPath}");
        }

        private async Task DownloadFile(HttpClient client, string fileName, string downloadUrl, string downloadPath)
        {
            LogMessage($"Downloading: {fileName}");

            var fileBytes = await client.GetByteArrayAsync(downloadUrl);
            var filePath = Path.Combine(downloadPath, fileName);

            // Use synchronous version for .NET Framework compatibility
            File.WriteAllBytes(filePath, fileBytes);
            LogMessage($"Downloaded: {fileName} ({fileBytes.Length / 1024} KB)");
        }
        #endregion

        #region JSON Parsing
        private string ExtractJsonValue(string json, string key)
        {
            if (string.IsNullOrEmpty(json) || string.IsNullOrEmpty(key))
                return string.Empty;

            var searchPattern = $"\"{key}\":\"";
            var startIndex = json.IndexOf(searchPattern, StringComparison.Ordinal);
            if (startIndex == -1) return string.Empty;

            startIndex += searchPattern.Length;
            var endIndex = json.IndexOf("\"", startIndex, StringComparison.Ordinal);

            return endIndex > startIndex ? json.Substring(startIndex, endIndex - startIndex) : string.Empty;
        }

        private string ExtractJsonSection(string json, string sectionName)
        {
            if (string.IsNullOrEmpty(json) || string.IsNullOrEmpty(sectionName))
                return string.Empty;

            var searchPattern = $"\"{sectionName}\":[";
            var startIndex = json.IndexOf(searchPattern, StringComparison.Ordinal);
            if (startIndex == -1) return string.Empty;

            startIndex += searchPattern.Length - 1;
            var bracketCount = 0;
            var endIndex = startIndex;

            for (int i = startIndex; i < json.Length; i++)
            {
                if (json[i] == '[') bracketCount++;
                if (json[i] == ']') bracketCount--;
                if (bracketCount == 0)
                {
                    endIndex = i;
                    break;
                }
            }

            return json.Substring(startIndex, endIndex - startIndex + 1);
        }

        private List<(string fileName, string downloadUrl)> ExtractDownloadUrls(string assetsJson)
        {
            var results = new List<(string, string)>();

            if (string.IsNullOrEmpty(assetsJson))
                return results;

            var parts = assetsJson.Split(new[] { "}," }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var part in parts)
            {
                var name = ExtractJsonValue(part, "name");
                var downloadUrl = ExtractJsonValue(part, "browser_download_url");

                if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(downloadUrl))
                {
                    results.Add((name, downloadUrl));
                }
            }

            return results;
        }
        #endregion

        #region File Extraction
        private async Task ExtractDownloadedFiles()
        {
            LogMessage("=== Extracting Downloaded Files ===");

            try
            {
                var downloadPath = GetTempPath(TEMP_DOWNLOAD_FOLDER);
                var extractPath = GetTempPath(TEMP_EXTRACT_FOLDER);

                Directory.CreateDirectory(extractPath);

                await ExtractZipFiles(downloadPath, extractPath);
                CopyNonZipFiles(downloadPath, extractPath);

                LogMessage($"All files extracted to: {extractPath}");
            }
            catch (Exception ex)
            {
                throw new ExtractionException($"Failed to extract files: {ex.Message}", ex);
            }

            LogMessage("");
        }

        private async Task ExtractZipFiles(string downloadPath, string extractPath)
        {
            var zipFiles = Directory.GetFiles(downloadPath, "*.zip");

            foreach (var zipFile in zipFiles)
            {
                LogMessage($"Extracting: {Path.GetFileName(zipFile)}");

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
            var entryPath = Path.Combine(extractPath, entry.FullName);
            var directory = Path.GetDirectoryName(entryPath);

            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            entry.ExtractToFile(entryPath, true);
            LogMessage($"   Extracted: {entry.Name}");

            await Task.Delay(50); // Visual feedback
        }

        private void CopyNonZipFiles(string downloadPath, string extractPath)
        {
            var otherFiles = Directory.GetFiles(downloadPath)
                .Where(f => !f.EndsWith(".zip", StringComparison.OrdinalIgnoreCase));

            foreach (var file in otherFiles)
            {
                var fileName = Path.GetFileName(file);
                var destPath = Path.Combine(extractPath, fileName);
                File.Copy(file, destPath, true);
                LogMessage($"Copied: {fileName}");
            }
        }
        #endregion

        #region DLL Registration
        private async Task RegisterDLLs()
        {
            LogMessage("=== Registering COM Components ===");

            try
            {
                var extractPath = GetTempPath(TEMP_EXTRACT_FOLDER);

                await CopyDLLsToInstallPath(extractPath);

                // Find the AresLicenseValidator DLL (flexible search)
                var validatorDll = FindAresLicenseValidatorDll();

                if (string.IsNullOrEmpty(validatorDll) || !File.Exists(validatorDll))
                {
                    // Log what we found in the Rsc folder for debugging
                    LogMessage("DEBUG: Files in Rsc folder:");
                    var rscFiles = Directory.GetFiles(DLL_PATH);
                    foreach (var file in rscFiles)
                    {
                        LogMessage($"  - {Path.GetFileName(file)}");
                    }

                    throw new FileNotFoundException("AresLicenseValidator.dll not found after copying to Rsc folder");
                }

                await RegisterSingleDLL(validatorDll);
                LogMessage("COM registration completed");
            }
            catch (Exception ex)
            {
                throw new RegistrationException($"Failed to register DLLs: {ex.Message}", ex);
            }

            LogMessage("");
        }

        private string FindAresLicenseValidatorDll()
        {
            try
            {
                // Search for any AresLicenseValidator DLL (flexible)
                var searchPatterns = new[]
                {
            "AresLicenseValidator.dll",      // Without version
            "AresLicenseValidator-*.dll"     // With version pattern
        };

                foreach (var pattern in searchPatterns)
                {
                    var files = Directory.GetFiles(DLL_PATH, pattern);
                    if (files.Length > 0)
                    {
                        LogMessage($"Found validator DLL: {Path.GetFileName(files[0])}");
                        return files[0];
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error searching for DLL: {ex.Message}");
            }

            return null;
        }

        private async Task CopyDLLsToInstallPath(string sourcePath)
        {
            // Find all DLL files in source
            var dllFiles = Directory.GetFiles(sourcePath, "*.dll");

            LogMessage($"Found {dllFiles.Length} DLL file(s) in download");

            if (dllFiles.Length == 0)
            {
                LogMessage("WARNING: No DLL files found in download");
                return;
            }

            foreach (var sourceDll in dllFiles)
            {
                await CopySingleDLL(sourcePath, Path.GetFileName(sourceDll));
            }
        }

        private async Task CopySingleDLL(string sourcePath, string dllFileName)
        {
            var sourceDll = Path.Combine(sourcePath, dllFileName);

            if (!File.Exists(sourceDll))
            {
                LogMessage($"Not found: {dllFileName}");
                return;
            }

            // Extract base name without version
            // e.g., "AresLicenseValidator-1.0.0.dll" or "AresLicenseValidator.dll" 
            // both become "AresLicenseValidator"
            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(dllFileName);
            var dllBaseName = fileNameWithoutExt;

            var lastDashIndex = fileNameWithoutExt.LastIndexOf('-');
            if (lastDashIndex > 0)
            {
                dllBaseName = fileNameWithoutExt.Substring(0, lastDashIndex);
            }

            // Check if same version already exists (only if source has version)
            var existingDll = FindExistingDllWithBaseName(dllBaseName);

            if (!string.IsNullOrEmpty(existingDll))
            {
                // Only check version if both files have versions in their names
                if (lastDashIndex > 0 && IsSameVersion(sourceDll, existingDll))
                {
                    LogMessage($"Same version already installed: {Path.GetFileName(existingDll)} - Skipping");
                    return;
                }

                // Backup old version
                var backupPath = Path.Combine(INSTALL_PATH, "Backup",
                    $"{Path.GetFileName(existingDll)}.backup_{DateTime.Now:yyyyMMdd_HHmmss}");

                File.Move(existingDll, backupPath);
                LogMessage($"Backed up old version: {Path.GetFileName(existingDll)}");

                // Also backup the TLB file if it exists
                var existingTlb = Path.ChangeExtension(existingDll, ".tlb");
                if (File.Exists(existingTlb))
                {
                    var tlbBackupPath = Path.Combine(INSTALL_PATH, "Backup",
                        $"{Path.GetFileName(existingTlb)}.backup_{DateTime.Now:yyyyMMdd_HHmmss}");
                    File.Move(existingTlb, tlbBackupPath);
                    LogMessage($"Backed up old TLB: {Path.GetFileName(existingTlb)}");
                }
            }

            // Copy new version
            var targetDll = Path.Combine(DLL_PATH, dllFileName);
            File.Copy(sourceDll, targetDll, true);
            LogMessage($"Copied: {dllFileName} to Rsc folder");

            // Also copy TLB if it exists in source
            var sourceTlb = Path.ChangeExtension(sourceDll, ".tlb");
            if (File.Exists(sourceTlb))
            {
                var targetTlb = Path.Combine(DLL_PATH, Path.GetFileName(sourceTlb));
                File.Copy(sourceTlb, targetTlb, true);
                LogMessage($"Copied: {Path.GetFileName(sourceTlb)} to Rsc folder");
            }

            await Task.Delay(100);
        }

        private async Task RegisterSingleDLL(string dllPath)
        {
            LogMessage($"Registering: {Path.GetFileName(dllPath)}");

            var regasmPath = FindRegAsmPath();
            if (string.IsNullOrEmpty(regasmPath))
            {
                throw new FileNotFoundException("RegAsm.exe not found. Please install .NET Framework Developer Pack.");
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
                    throw new RegistrationException($"Failed to execute RegAsm: {ex.Message}", ex);
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
                LogMessage("DLL registered successfully");
                if (!string.IsNullOrWhiteSpace(output))
                {
                    LogMessage($"RegAsm output: {output.Trim()}");
                }
            }
            else
            {
                throw new RegistrationException($"RegAsm failed (Exit code: {exitCode}): {error}");
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
            LogMessage("=== Installing ARES Project ===");

            try
            {
                var extractPath = GetTempPath(TEMP_EXTRACT_FOLDER);

                await CopyMVBAFile(extractPath);

                await Task.Delay(500);
            }
            catch (Exception ex)
            {
                throw new InstallationException($"Failed to copy ARES project: {ex.Message}", ex);
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
                LogMessage($"Copied ARES.mvba to: {mvbaTarget}");
            }
            else
            {
                LogMessage("ARES.mvba not found in download");
            }

            await Task.Delay(100);
        }
        #endregion

        #region Cleanup
        private async Task CleanupTemporaryFiles()
        {
            LogMessage("=== Cleaning Up ===");

            try
            {
                var tempPaths = new[]
                {
                    GetTempPath(TEMP_DOWNLOAD_FOLDER),
                    GetTempPath(TEMP_EXTRACT_FOLDER)
                };

                foreach (var tempPath in tempPaths)
                {
                    if (Directory.Exists(tempPath))
                    {
                        Directory.Delete(tempPath, true);
                        LogMessage($"Cleaned: {tempPath}");
                        await Task.Delay(100);
                    }
                }

                LogMessage("Cleanup completed");
            }
            catch (Exception ex)
            {
                LogMessage($"Cleanup warning: {ex.Message}");
            }

            LogMessage("");
        }
        #endregion

        #region Utility Methods
        private string GetTempPath(string folderName)
        {
            return Path.Combine(Path.GetTempPath(), folderName);
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
            var logEntry = $"{DateTime.Now:HH:mm:ss} - {message}\n";

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
                // Prevent closing while installation is in progress
                var message = currentLanguage == "EN"
                    ? "Installation in progress. Please wait..."
                    : "Installation en cours. Veuillez patienter...";

                MessageBox.Show(message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
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