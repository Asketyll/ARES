using System;
using System.Windows.Forms;

namespace AresInstaller
{
    public partial class LanguageSelectionForm : Form
    {
        public string SelectedLanguage { get; private set; }

        private Button btnEnglish;
        private Button btnFrench;
        private Label lblTitle;

        public LanguageSelectionForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Language Selection / Sélection de la langue";
            this.Size = new System.Drawing.Size(400, 200);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Title Label
            lblTitle = new Label
            {
                Text = "Choose installation language\nChoisissez la langue d'installation",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(340, 50),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold)
            };
            this.Controls.Add(lblTitle);

            // English Button
            btnEnglish = new Button
            {
                Text = "English",
                Location = new System.Drawing.Point(50, 90),
                Size = new System.Drawing.Size(120, 40),
                Font = new System.Drawing.Font("Segoe UI", 10)
            };
            btnEnglish.Click += BtnEnglish_Click;
            this.Controls.Add(btnEnglish);

            // French Button
            btnFrench = new Button
            {
                Text = "Français",
                Location = new System.Drawing.Point(210, 90),
                Size = new System.Drawing.Size(120, 40),
                Font = new System.Drawing.Font("Segoe UI", 10)
            };
            btnFrench.Click += BtnFrench_Click;
            this.Controls.Add(btnFrench);
        }

        private void BtnEnglish_Click(object sender, EventArgs e)
        {
            SelectedLanguage = "EN";
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BtnFrench_Click(object sender, EventArgs e)
        {
            SelectedLanguage = "FR";
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}