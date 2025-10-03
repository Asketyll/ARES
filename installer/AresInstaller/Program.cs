using System;
using System.Diagnostics;
using System.Security.Principal;
using System.Windows.Forms;

namespace AresInstaller
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            // Check if running as administrator
            if (!IsRunAsAdministrator())
            {
                // Restart as administrator
                RestartAsAdministrator();
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Show language selection first
            using (var langForm = new LanguageSelectionForm())
            {
                if (langForm.ShowDialog() == DialogResult.OK)
                {
                    Application.Run(new AresInstallerForm(langForm.SelectedLanguage));
                }
            }
        }

        private static bool IsRunAsAdministrator()
        {
            try
            {
                WindowsIdentity identity = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch
            {
                return false;
            }
        }

        private static void RestartAsAdministrator()
        {
            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    UseShellExecute = true,
                    WorkingDirectory = Environment.CurrentDirectory,
                    FileName = Application.ExecutablePath,
                    Verb = "runas"
                };

                Process.Start(startInfo);
            }
            catch
            {
                // User cancelled UAC prompt
                MessageBox.Show(
                    "This installer requires administrator privileges to run.\n\nCet installateur nécessite des privilèges administrateur pour s'exécuter.",
                    "Administrator Required / Administrateur requis",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
            }

            Application.Exit();
        }
    }
}