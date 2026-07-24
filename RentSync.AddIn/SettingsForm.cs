using System;
using System.Drawing;
using System.Windows.Forms;

namespace RentSync.AddIn
{
    /// <summary>
    /// Minimal settings dialog: stores the API token in Windows Credential
    /// Manager. Deliberately plain WinForms for now — the DevExpress-styled
    /// version replaces this in the next iteration without touching any
    /// other file (the form is self-contained).
    /// </summary>
    public sealed class SettingsForm : Form
    {
        private readonly TextBox _token = new TextBox
            { UseSystemPasswordChar = true, Width = 320 };

        public SettingsForm()
        {
            Text = "RentSync Settings";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = MaximizeBox = false;
            ClientSize = new Size(380, 120);

            var label = new Label
            {
                Text = "API token (stored in Windows Credential Manager):",
                AutoSize = true, Left = 20, Top = 15
            };
            _token.Left = 20; _token.Top = 40;

            var save = new Button { Text = "Save", Left = 20, Top = 75, Width = 90 };
            save.Click += OnSave;
            var cancel = new Button { Text = "Cancel", Left = 120, Top = 75, Width = 90,
                DialogResult = DialogResult.Cancel };

            Controls.AddRange(new Control[] { label, _token, save, cancel });
            AcceptButton = save;
            CancelButton = cancel;
        }

        private void OnSave(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(_token.Text))
                WindowsCredentialTokenProvider.WriteCredential(
                    WindowsCredentialTokenProvider.TargetName, _token.Text.Trim());
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
