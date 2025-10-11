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

            // Check administrator privileges (silent check, it's always true due to manifest)
            if (!IsRunningAsAdministrator())
            {
                LogMessage(Translations.Get("NotRunningAsAdmin", currentLanguage));
            }

            // Check .NET Framework
            if (IsDotNetFrameworkInstalled())
            {
                LogMessage(Translations.Get("DotNetAvailable", currentLanguage));
            }
            else
            {
                throw new InvalidOperationException(Translations.Get("DotNetRequired", currentLanguage));
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
                    var releaseInfo = await GetLatestReleaseInfo(client);

                    LogMessage(Translations.Get("ParsingReleaseAssets", currentLanguage));
                    await DownloadReleaseAssets(client, releaseInfo);
                }
                catch (Exception ex)
                {
                    // DEBUG messages stay in English
                    LogMessage($"DOWNLOAD ERROR: {ex.GetType().Name}");
                    LogMessage($"Message: {ex.Message}");
                    if (ex.InnerException != null)
                    {
                        LogMessage($"Inner Exception: {ex.InnerException.Message}");
                    }
                    LogMessage($"Stack Trace: {ex.StackTrace}");
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

                // Parse with Newtonsoft.Json to extract tag name
                var release = JObject.Parse(releaseResponse);
                var tagName = release["tag_name"]?.ToString();

                if (string.IsNullOrEmpty(tagName))
                {
                    LogMessage("WARNING: Could not find tag_name in response"); // DEBUG
                }
                else
                {
                    LogMessage(Translations.Format("LatestVersion", currentLanguage, tagName));
                }

                return releaseResponse;
            }
            catch (HttpRequestException httpEx)
            {
                LogMessage($"HTTP ERROR: {httpEx.Message}"); // DEBUG
                throw;
            }
            catch (Newtonsoft.Json.JsonException jsonEx)
            {
                LogMessage($"JSON PARSE ERROR: {jsonEx.Message}"); // DEBUG
                throw;
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: {ex.Message}"); // DEBUG
                throw;
            }
        }

        private async Task DownloadReleaseAssets(HttpClient client, string releaseResponse)
        {
            try
            {
                // Parse JSON with Newtonsoft.Json
                var release = JObject.Parse(releaseResponse);
                var assets = release["assets"] as JArray;

                if (assets == null || assets.Count == 0)
                {
                    LogMessage("ERROR: No assets found in release!"); // DEBUG
                    LogMessage($"Full release response: {releaseResponse}"); // DEBUG
                    throw new DownloadException(Translations.Get("NoAssetsFound", currentLanguage));
                }

                LogMessage(Translations.Format("FoundAssets", currentLanguage, assets.Count));

                var downloadPath = GetTempPath(TEMP_DOWNLOAD_FOLDER);
                Directory.CreateDirectory(downloadPath);

                if (!Directory.Exists(downloadPath))
                {
                    throw new DownloadException(Translations.Format("FailedCreateDownloadDir", currentLanguage, downloadPath));
                }

                // Download each asset
                foreach (var asset in assets)
                {
                    var fileName = asset["name"]?.ToString();
                    var downloadUrl = asset["browser_download_url"]?.ToString();

                    if (!string.IsNullOrEmpty(fileName) && !string.IsNullOrEmpty(downloadUrl))
                    {
                        try
                        {
                            await DownloadFile(client, fileName, downloadUrl, downloadPath);
                        }
                        catch (Exception ex)
                        {
                            LogMessage($"ERROR downloading {fileName}: {ex.Message}"); // DEBUG
                            throw;
                        }
                    }
                    else
                    {
                        LogMessage($"WARNING: Skipped asset with missing name or URL"); // DEBUG
                    }
                }

                // Summary: show total files downloaded
                var downloadedFiles = Directory.GetFiles(downloadPath);
                LogMessage(Translations.Format("DownloadComplete", currentLanguage, downloadedFiles.Length));
            }
            catch (Newtonsoft.Json.JsonException jsonEx)
            {
                LogMessage($"JSON PARSE ERROR: {jsonEx.Message}"); // DEBUG
                throw new DownloadException(Translations.Format("FailedParseJSON", currentLanguage, jsonEx.Message), jsonEx);
            }
            catch (DownloadException)
            {
                // Re-throw DownloadException as-is
                throw;
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: {ex.Message}"); // DEBUG
                throw new DownloadException(Translations.Format("FailedDownloadAssets", currentLanguage, ex.Message), ex);
            }
        }

        private async Task DownloadFile(HttpClient client, string fileName, string downloadUrl, string downloadPath)
        {
            LogMessage(Translations.Format("DownloadingFile", currentLanguage, fileName));

            try
            {
                var fileBytes = await client.GetByteArrayAsync(downloadUrl);
                var filePath = Path.Combine(downloadPath, fileName);
                File.WriteAllBytes(filePath, fileBytes);

                // Verify file was written (silent check)
                if (!File.Exists(filePath))
                {
                    throw new IOException($"File was not created: {filePath}"); // DEBUG message
                }
            }
            catch (HttpRequestException httpEx)
            {
                LogMessage($"HTTP ERROR: {httpEx.Message}"); // DEBUG
                throw;
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: {ex.GetType().Name} - {ex.Message}"); // DEBUG
                throw;
            }
        }
        #endregion

        #region File Extraction
        private async Task ExtractDownloadedFiles()
        {
            LogMessage(Translations.Get("ExtractingFiles", currentLanguage));

            try
            {
                var downloadPath = GetTempPath(TEMP_DOWNLOAD_FOLDER);
                var extractPath = GetTempPath(TEMP_EXTRACT_FOLDER);

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
            var entryPath = Path.Combine(extractPath, entry.FullName);
            var directory = Path.GetDirectoryName(entryPath);

            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            entry.ExtractToFile(entryPath, true);
            LogMessage(Translations.Format("Extracted", currentLanguage, entry.Name));

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
                var extractPath = GetTempPath(TEMP_EXTRACT_FOLDER);

                await CopyDLLsToInstallPath(extractPath);

                // DEBUG: List all files in Rsc folder
                LogMessage("DEBUG: Listing all files in Rsc folder:");
                var rscFiles = Directory.GetFiles(DLL_PATH);
                LogMessage($"  Total files: {rscFiles.Length}");
                foreach (var file in rscFiles)
                {
                    LogMessage($"  - {Path.GetFileName(file)}");
                }

                // Find the AresLicenseValidator DLL (flexible search)
                LogMessage(Translations.Get("SearchingValidator", currentLanguage));
                var validatorDll = FindAresLicenseValidatorDll();

                if (string.IsNullOrEmpty(validatorDll))
                {
                    LogMessage("ERROR: FindAresLicenseValidatorDll() returned null or empty"); // DEBUG
                    throw new FileNotFoundException(Translations.Get("ValidatorNotFound", currentLanguage));
                }

                if (!File.Exists(validatorDll))
                {
                    LogMessage($"ERROR: File.Exists() returned false for: {validatorDll}"); // DEBUG
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
                LogMessage($"Searching in directory: {DLL_PATH}"); // DEBUG

                // Check if directory exists
                if (!Directory.Exists(DLL_PATH))
                {
                    LogMessage($"ERROR: Directory does not exist: {DLL_PATH}"); // DEBUG
                    return null;
                }

                // Search for any AresLicenseValidator DLL (flexible)
                var searchPatterns = new[]
                {
                    "AresLicenseValidator.dll",      // Without version
                    "AresLicenseValidator-*.dll"     // With version pattern
                };

                foreach (var pattern in searchPatterns)
                {
                    LogMessage($"Searching with pattern: {pattern}"); // DEBUG
                    var files = Directory.GetFiles(DLL_PATH, pattern);
                    LogMessage($"  Found {files.Length} file(s)"); // DEBUG

                    if (files.Length > 0)
                    {
                        var foundFile = files[0];
                        LogMessage($"  Using: {Path.GetFileName(foundFile)}"); // DEBUG
                        LogMessage($"  Full path: {foundFile}"); // DEBUG
                        LogMessage($"  File exists: {File.Exists(foundFile)}"); // DEBUG
                        return foundFile;
                    }
                }

                LogMessage("No matching DLL found"); // DEBUG
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR in FindAresLicenseValidatorDll: {ex.Message}"); // DEBUG
                LogMessage($"Stack trace: {ex.StackTrace}"); // DEBUG
            }

            return null;
        }

        private async Task CopyDLLsToInstallPath(string sourcePath)
        {
            // Find all DLL files in source (they may have version suffixes)
            var dllFiles = Directory.GetFiles(sourcePath, "*.dll");

            foreach (var sourceDll in dllFiles)
            {
                await CopySingleDLL(sourcePath, Path.GetFileName(sourceDll));
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

            // Extract base name without version
            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(dllFileName);
            var dllBaseName = fileNameWithoutExt;

            var lastDashIndex = fileNameWithoutExt.LastIndexOf('-');
            if (lastDashIndex > 0)
            {
                dllBaseName = fileNameWithoutExt.Substring(0, lastDashIndex);
            }

            // Check if same version already exists
            var existingDll = FindExistingDllWithBaseName(dllBaseName);

            if (!string.IsNullOrEmpty(existingDll))
            {
                // Only check version if both files have versions in their names
                if (lastDashIndex > 0 && IsSameVersion(sourceDll, existingDll))
                {
                    LogMessage(Translations.Get("SameVersionInstalled", currentLanguage));
                    return;
                }

                // Backup old version
                var backupPath = Path.Combine(INSTALL_PATH, "Backup",
                    $"{Path.GetFileName(existingDll)}.backup_{DateTime.Now:yyyyMMdd_HHmmss}");

                File.Move(existingDll, backupPath);
                LogMessage(Translations.Format("BackedUpOldVersion", currentLanguage, Path.GetFileName(backupPath)));

                // Also backup the TLB file if it exists
                var existingTlb = Path.ChangeExtension(existingDll, ".tlb");
                if (File.Exists(existingTlb))
                {
                    var tlbBackupPath = Path.Combine(INSTALL_PATH, "Backup",
                        $"{Path.GetFileName(existingTlb)}.backup_{DateTime.Now:yyyyMMdd_HHmmss}");
                    File.Move(existingTlb, tlbBackupPath);
                    LogMessage(Translations.Format("BackedUpOldTLB", currentLanguage, Path.GetFileName(tlbBackupPath)));
                }
            }

            // Copy new version
            var targetDll = Path.Combine(DLL_PATH, dllFileName);
            File.Copy(sourceDll, targetDll, true);

            // Verify copy
            if (File.Exists(targetDll))
            {
                LogMessage(Translations.Format("CopiedDLL", currentLanguage, dllFileName, new FileInfo(targetDll).Length));
            }
            else
            {
                throw new IOException(Translations.Format("FailedCopyDLL", currentLanguage, targetDll));
            }

            // Also copy TLB if it exists in source (silent)
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
                // RegAsm output is too verbose and not useful for end users
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
                var extractPath = GetTempPath(TEMP_EXTRACT_FOLDER);
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
                        await Task.Delay(100);
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