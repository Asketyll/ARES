using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AresInstaller
{
    public partial class BentleyProductSelectionForm : Form
    {
        private const string BENTLEY_REGISTRY_PATH = @"SOFTWARE\Bentley\Installed_Products";
        private static readonly string[] VALID_PRODUCTS = { "MapPowerView", "Microstation" };
        private const string ARES_MVBA_PATH = "c:/ares/ares.mvba";
        private const string AUTOLOAD_LINE = "MS_VBAAUTOLOADPROJECTS > " + ARES_MVBA_PATH;

        private ComboBox productComboBox;
        private Label titleLabel;
        private Button nextButton;
        private Button cancelButton;

        private List<BentleyProduct> bentleyProducts;
        private string currentLanguage;

        public string SelectedConfigurationPath { get; private set; }
        public BentleyProduct SelectedProduct { get; private set; }

        public BentleyProductSelectionForm(string language)
        {
            currentLanguage = language;
            bentleyProducts = new List<BentleyProduct>();
            InitializeComponent();
            LoadBentleyProducts();
            ApplyTranslations();
        }

        private void InitializeComponent()
        {
            this.Size = new System.Drawing.Size(600, 200);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Title Label
            titleLabel = new Label
            {
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(540, 30),
                Font = new System.Drawing.Font("Segoe UI", 11, System.Drawing.FontStyle.Bold)
            };
            this.Controls.Add(titleLabel);

            // Product ComboBox Label
            Label comboLabel = new Label
            {
                Location = new System.Drawing.Point(20, 60),
                Size = new System.Drawing.Size(540, 20),
                Text = "Select Bentley Product / Sélectionnez un produit Bentley:"
            };
            this.Controls.Add(comboLabel);

            // Product ComboBox
            productComboBox = new ComboBox
            {
                Location = new System.Drawing.Point(20, 85),
                Size = new System.Drawing.Size(540, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new System.Drawing.Font("Segoe UI", 10)
            };
            productComboBox.SelectedIndexChanged += ProductComboBox_SelectedIndexChanged;
            this.Controls.Add(productComboBox);

            // Next Button
            nextButton = new Button
            {
                Location = new System.Drawing.Point(350, 125),
                Size = new System.Drawing.Size(100, 35),
                Enabled = false,
                Font = new System.Drawing.Font("Segoe UI", 10)
            };
            nextButton.Click += NextButton_Click;
            this.Controls.Add(nextButton);

            // Cancel Button
            cancelButton = new Button
            {
                Location = new System.Drawing.Point(460, 125),
                Size = new System.Drawing.Size(100, 35),
                Font = new System.Drawing.Font("Segoe UI", 10)
            };
            cancelButton.Click += CancelButton_Click;
            this.Controls.Add(cancelButton);
        }

        private void LoadBentleyProducts()
        {
            try
            {
                RegistryKey baseKey = null;

                // Try to open registry in 64-bit view first, then 32-bit view
                try
                {
                    baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64)
                        .OpenSubKey(BENTLEY_REGISTRY_PATH);
                }
                catch
                {
                    // Ignore and try 32-bit view
                }

                if (baseKey == null)
                {
                    try
                    {
                        baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
                            .OpenSubKey(BENTLEY_REGISTRY_PATH);
                    }
                    catch
                    {
                        // Ignore
                    }
                }

                using (baseKey)
                {
                    if (baseKey == null)
                    {
                        MessageBox.Show(
                            Translations.Get("NoBentleyProducts", currentLanguage),
                            Translations.Get("ProductSelection", currentLanguage),
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        );
                        return;
                    }

                    string[] subKeyNames = baseKey.GetSubKeyNames();

                    foreach (string subKeyName in subKeyNames)
                    {
                        using (RegistryKey productKey = baseKey.OpenSubKey(subKeyName))
                        {
                            if (productKey != null)
                            {
                                string productName = productKey.GetValue("ProductName") as string;

                                // Check if product is valid (MapPowerView or Microstation)
                                if (!string.IsNullOrEmpty(productName) &&
                                    Array.Exists(VALID_PRODUCTS, p => p.Equals(productName, StringComparison.OrdinalIgnoreCase)))
                                {
                                    string displayName = productKey.GetValue("DisplayProductName") as string;
                                    string version = productKey.GetValue("Version") as string;

                                    // Try to find configuration path with flexible key name matching
                                    string configPath = null;
                                    string[] valueNames = productKey.GetValueNames();

                                    foreach (string valueName in valueNames)
                                    {
                                        if (valueName.StartsWith("Configuration", StringComparison.OrdinalIgnoreCase))
                                        {
                                            configPath = productKey.GetValue(valueName) as string;
                                            if (!string.IsNullOrEmpty(configPath))
                                                break;
                                        }
                                    }

                                    if (!string.IsNullOrEmpty(displayName) && !string.IsNullOrEmpty(version))
                                    {
                                        bentleyProducts.Add(new BentleyProduct
                                        {
                                            DisplayName = displayName,
                                            Version = version,
                                            ConfigurationPath = configPath ?? "",
                                            ProductName = productName
                                        });
                                    }
                                }
                            }
                        }
                    }
                }

                // Populate ComboBox
                if (bentleyProducts.Count > 0)
                {
                    productComboBox.DataSource = bentleyProducts;
                    productComboBox.DisplayMember = "ToString";
                }
                else
                {
                    MessageBox.Show(
                        Translations.Get("NoValidBentleyProducts", currentLanguage),
                        Translations.Get("ProductSelection", currentLanguage),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    Translations.Format("FailedLoadProducts", currentLanguage, ex.Message),
                    Translations.Get("ProductSelectionError", currentLanguage),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        private void ProductComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (productComboBox.SelectedItem is BentleyProduct selectedProduct)
            {
                SelectedProduct = selectedProduct;
                SelectedConfigurationPath = selectedProduct.ConfigurationPath;
                nextButton.Enabled = true;
            }
        }

        private async void NextButton_Click(object sender, EventArgs e)
        {
            if (SelectedProduct == null)
                return;

            nextButton.Enabled = false;
            cancelButton.Enabled = false;

            try
            {
                // Hide the form
                this.Hide();

                // Configure ARES autoload
                await ConfigureAresAutoload();

                // Show success message
                MessageBox.Show(
                    Translations.Get("ConfigurationSuccess", currentLanguage),
                    Translations.Get("Configuration", currentLanguage),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                this.Show();
                MessageBox.Show(
                    Translations.Format("ConfigurationError", currentLanguage, ex.Message),
                    Translations.Get("Configuration", currentLanguage),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );

                nextButton.Enabled = true;
                cancelButton.Enabled = true;
            }
        }

        private async System.Threading.Tasks.Task ConfigureAresAutoload()
        {
            await System.Threading.Tasks.Task.Run(() =>
            {
                // Get LocalAppData path
                string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

                // Build Bentley product path
                string bentleyProductPath = Path.Combine(localAppData, "Bentley", SelectedProduct.ProductName);

                if (!Directory.Exists(bentleyProductPath))
                {
                    throw new DirectoryNotFoundException(
                        Translations.Format("BentleyProductPathNotFound", currentLanguage, bentleyProductPath)
                    );
                }

                // Find all Personal.ucf files
                List<string> personalUcfFiles = FindPersonalUcfFiles(bentleyProductPath);

                if (personalUcfFiles.Count == 0)
                {
                    throw new FileNotFoundException(
                        Translations.Format("NoPersonalUcfFound", currentLanguage, bentleyProductPath)
                    );
                }

                // Process each Personal.ucf file
                int filesModified = 0;
                foreach (string ucfFile in personalUcfFiles)
                {
                    if (ProcessPersonalUcfFile(ucfFile))
                    {
                        filesModified++;
                    }
                }

                if (filesModified == 0)
                {
                    // All files already configured
                    MessageBox.Show(
                        Translations.Get("AlreadyConfigured", currentLanguage),
                        Translations.Get("Configuration", currentLanguage),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            });
        }

        private List<string> FindPersonalUcfFiles(string rootPath)
        {
            List<string> ucfFiles = new List<string>();

            try
            {
                // Get all subdirectories
                string[] subdirectories = Directory.GetDirectories(rootPath);

                foreach (string subdirectory in subdirectories)
                {
                    // Look for prefs folder in each subdirectory
                    string prefsPath = Path.Combine(subdirectory, "prefs");

                    if (Directory.Exists(prefsPath))
                    {
                        string personalUcfPath = Path.Combine(prefsPath, "Personal.ucf");

                        if (File.Exists(personalUcfPath))
                        {
                            ucfFiles.Add(personalUcfPath);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(
                    Translations.Format("ErrorSearchingUcfFiles", currentLanguage, ex.Message),
                    ex
                );
            }

            return ucfFiles;
        }

        private bool ProcessPersonalUcfFile(string ucfFilePath)
        {
            try
            {
                // Read all lines
                string[] lines = File.ReadAllLines(ucfFilePath, Encoding.UTF8);

                // Check if autoload line already exists
                bool lineExists = false;
                foreach (string line in lines)
                {
                    if (line.Trim().Equals(AUTOLOAD_LINE, StringComparison.OrdinalIgnoreCase))
                    {
                        lineExists = true;
                        break;
                    }
                }

                if (!lineExists)
                {
                    // Add the autoload line
                    List<string> newLines = new List<string>(lines);
                    newLines.Add(AUTOLOAD_LINE);

                    // Write back to file
                    File.WriteAllLines(ucfFilePath, newLines, Encoding.UTF8);
                    return true; // File was modified
                }

                return false; // File was not modified (line already exists)
            }
            catch (Exception ex)
            {
                throw new Exception(
                    Translations.Format("ErrorProcessingUcfFile", currentLanguage, ucfFilePath, ex.Message),
                    ex
                );
            }
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void ApplyTranslations()
        {
            this.Text = Translations.Get("ProductSelection", currentLanguage);
            titleLabel.Text = Translations.Get("SelectBentleyProduct", currentLanguage);
            nextButton.Text = Translations.Get("NextButton", currentLanguage);
            cancelButton.Text = Translations.Get("CancelButton", currentLanguage);
        }
    }

    public class BentleyProduct
    {
        public string DisplayName { get; set; }
        public string Version { get; set; }
        public string ConfigurationPath { get; set; }
        public string ProductName { get; set; }

        public override string ToString()
        {
            return $"{DisplayName} - {Version}";
        }
    }
}