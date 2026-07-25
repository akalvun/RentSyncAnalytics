using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Extensions.DependencyInjection;
using RentSync.Core.Services;
using RentSync.Core.Telemetry;

namespace RentSync.AddIn
{
    /// <summary>
    /// Settings dialog for the add-in.
    ///
    /// Two constraints shape this form, and both come from living inside
    /// Excel's process rather than owning it:
    ///
    /// 1. DPI. An add-in cannot change the DPI awareness mode of its host,
    ///    so it has to survive whatever Excel was started with. Every
    ///    control is placed by TableLayoutPanel and the form auto-sizes to
    ///    its content — there is not one absolute coordinate in this file.
    ///    Hard-coded Left/Top is what makes add-in dialogs clip their own
    ///    buttons at 150% scaling.
    ///
    /// 2. Ownership. The dialog is shown with Excel's window as its owner
    ///    (see ExcelWindow.Owner), otherwise a modal dialog can end up
    ///    behind the workbook with Excel refusing input — the classic
    ///    "Excel has frozen" support call.
    ///
    /// The API token is written to Windows Credential Manager. A stored
    /// token is never read back into the text box: the dialog reports that
    /// one exists and offers to replace it. Displaying a secret just so the
    /// user can look at it is a risk with no matching benefit.
    /// </summary>
    public sealed class SettingsForm : Form
    {
        private readonly IServiceProvider _services;
        private readonly UserSettings _settings;

        private TextBox _token;
        private Button _showToken;
        private Label _tokenState;
        private RadioButton _demoMode;
        private RadioButton _liveMode;
        private TextBox _baseUrl;
        private NumericUpDown _freshness;
        private Label _restartHint;
        private Label _status;
        private LinkLabel _openLogs;
        private Button _test;
        private Button _save;
        private Button _cancel;

        private bool _busy;

        public SettingsForm(IServiceProvider services)
        {
            if (services == null) throw new ArgumentNullException("services");
            _services = services;
            _settings = UserSettings.Load();

            BuildUi();
            LoadValues();
        }

        // ---------------- layout ----------------

        private void BuildUi()
        {
            Text = "RentSync Settings";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(14);
            Font = SystemFonts.MessageBoxFont;

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                RowCount = 8
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            // --- row 0: API token ---
            _token = new TextBox
            {
                UseSystemPasswordChar = true,
                Width = 300,
                Anchor = AnchorStyles.Left | AnchorStyles.Right
            };
            _showToken = new Button { Text = "Show", AutoSize = true, Margin = new Padding(6, 0, 0, 0) };
            _showToken.Click += OnToggleTokenVisibility;

            var tokenRow = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                Dock = DockStyle.Fill,
                Margin = new Padding(0)
            };
            tokenRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tokenRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            tokenRow.Controls.Add(_token, 0, 0);
            tokenRow.Controls.Add(_showToken, 1, 0);

            root.Controls.Add(MakeLabel("API token:"), 0, 0);
            root.Controls.Add(tokenRow, 1, 0);

            // --- row 1: token state ---
            _tokenState = new Label
            {
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Margin = new Padding(3, 0, 3, 10)
            };
            root.Controls.Add(_tokenState, 1, 1);

            // --- row 2: data source ---
            _demoMode = new RadioButton { Text = "Demo data (bundled sample file)", AutoSize = true };
            _liveMode = new RadioButton { Text = "Live API", AutoSize = true, Margin = new Padding(16, 3, 3, 3) };
            _demoMode.CheckedChanged += OnRestartRelevantChange;
            _liveMode.CheckedChanged += OnRestartRelevantChange;

            var modeRow = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                Margin = new Padding(0),
                WrapContents = false
            };
            modeRow.Controls.Add(_demoMode);
            modeRow.Controls.Add(_liveMode);

            root.Controls.Add(MakeLabel("Data source:"), 0, 2);
            root.Controls.Add(modeRow, 1, 2);

            // --- row 3: base URL ---
            _baseUrl = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Width = 300 };
            _baseUrl.TextChanged += OnRestartRelevantChange;
            root.Controls.Add(MakeLabel("API base URL:"), 0, 3);
            root.Controls.Add(_baseUrl, 1, 3);

            // --- row 4: cache freshness ---
            _freshness = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 10080,
                Width = 70,
                Margin = new Padding(3, 3, 6, 3)
            };
            _freshness.ValueChanged += OnRestartRelevantChange;

            var freshnessRow = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                Margin = new Padding(0),
                WrapContents = false
            };
            freshnessRow.Controls.Add(_freshness);
            freshnessRow.Controls.Add(new Label
            {
                Text = "minutes before the cache is refetched",
                AutoSize = true,
                Margin = new Padding(0, 6, 0, 0)
            });

            root.Controls.Add(MakeLabel("Cache freshness:"), 0, 4);
            root.Controls.Add(freshnessRow, 1, 4);

            // --- row 5: restart hint ---
            _restartHint = new Label
            {
                AutoSize = true,
                Visible = false,
                ForeColor = SystemColors.GrayText,
                Text = "Data source and cache changes take effect the next time Excel starts.",
                Margin = new Padding(3, 10, 3, 0)
            };
            root.Controls.Add(_restartHint, 0, 5);
            root.SetColumnSpan(_restartHint, 2);

            // --- row 6: status ---
            _status = new Label
            {
                AutoSize = true,
                Margin = new Padding(3, 6, 3, 6),
                MaximumSize = new Size(430, 0)
            };
            root.Controls.Add(_status, 0, 6);
            root.SetColumnSpan(_status, 2);

            // --- row 7: buttons ---
            _openLogs = new LinkLabel
            {
                Text = "Open log folder",
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(3, 9, 3, 3)
            };
            _openLogs.LinkClicked += OnOpenLogs;

            _test = new Button { Text = "Test connection", AutoSize = true, Margin = new Padding(6, 0, 0, 0) };
            _test.Click += OnTestConnection;
            _save = new Button { Text = "Save", AutoSize = true, MinimumSize = new Size(84, 0), Margin = new Padding(6, 0, 0, 0) };
            _save.Click += OnSave;
            _cancel = new Button
            {
                Text = "Cancel",
                AutoSize = true,
                MinimumSize = new Size(84, 0),
                DialogResult = DialogResult.Cancel,
                Margin = new Padding(6, 0, 0, 0)
            };

            var buttons = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.RightToLeft,
                Anchor = AnchorStyles.Right,
                Margin = new Padding(0, 6, 0, 0),
                WrapContents = false
            };
            buttons.Controls.Add(_cancel);
            buttons.Controls.Add(_save);
            buttons.Controls.Add(_test);

            root.Controls.Add(_openLogs, 0, 7);
            root.Controls.Add(buttons, 1, 7);

            Controls.Add(root);
            AcceptButton = _save;
            CancelButton = _cancel;
        }

        private static Label MakeLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(3, 7, 12, 3)
            };
        }

        // ---------------- state ----------------

        private void LoadValues()
        {
            var stored = WindowsCredentialTokenProvider.ReadCredential(
                WindowsCredentialTokenProvider.TargetName);

            _tokenState.Text = string.IsNullOrEmpty(stored)
                ? "No token stored — the add-in runs anonymously."
                : "A token is stored. Type a new one to replace it.";

            _demoMode.Checked = _settings.UseDemoData;
            _liveMode.Checked = !_settings.UseDemoData;
            _baseUrl.Text = _settings.ApiBaseUrl;
            _freshness.Value = Math.Min(
                Math.Max(_settings.CacheFreshnessMinutes, (int)_freshness.Minimum),
                (int)_freshness.Maximum);

            UpdateEnabledState();
            _restartHint.Visible = false;
        }

        private void UpdateEnabledState()
        {
            _baseUrl.Enabled = _liveMode.Checked && !_busy;
            _token.Enabled = !_busy;
            _showToken.Enabled = !_busy;
            _freshness.Enabled = !_busy;
            _demoMode.Enabled = !_busy;
            _liveMode.Enabled = !_busy;
            _test.Enabled = !_busy;
            _save.Enabled = !_busy;
        }

        private void SetBusy(bool busy)
        {
            _busy = busy;
            UpdateEnabledState();
            Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
        }

        private void SetStatus(string text, bool isError)
        {
            _status.Text = text;
            _status.ForeColor = isError ? Color.Firebrick : SystemColors.ControlText;
        }

        // ---------------- handlers ----------------

        private void OnToggleTokenVisibility(object sender, EventArgs e)
        {
            _token.UseSystemPasswordChar = !_token.UseSystemPasswordChar;
            _showToken.Text = _token.UseSystemPasswordChar ? "Show" : "Hide";
        }

        private void OnRestartRelevantChange(object sender, EventArgs e)
        {
            UpdateEnabledState();
            _restartHint.Visible =
                _demoMode.Checked != _settings.UseDemoData ||
                (int)_freshness.Value != _settings.CacheFreshnessMinutes ||
                !string.Equals(_baseUrl.Text.Trim(), _settings.ApiBaseUrl, StringComparison.Ordinal);
        }

        private void OnOpenLogs(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                Directory.CreateDirectory(AppPaths.LogDirectory);
                Process.Start("explorer.exe", "\"" + AppPaths.LogDirectory + "\"");
            }
            catch (Exception ex)
            {
                SetStatus("Could not open the log folder: " + ex.Message, true);
            }
        }

        /// <summary>
        /// Exercises the real data path — the same service the ribbon calls —
        /// rather than a synthetic ping, so a green result means the feature
        /// works and not merely that a socket opened.
        ///
        /// Note that this tests the *running* configuration. Settings that
        /// need a restart are not applied to the container yet, which is
        /// exactly what the restart hint above is warning about.
        /// </summary>
        private async void OnTestConnection(object sender, EventArgs e)
        {
            SetBusy(true);
            SetStatus("Testing…", false);
            try
            {
                var rows = await _services.GetRequiredService<IRentRollDataService>()
                                          .GetRentRollAsync(forceRefresh: true);
                var count = 0;
                foreach (var row in rows) count++;

                SetStatus(
                    count > 0
                        ? string.Format("OK — {0} rows returned.", count)
                        : "Connected, but no rows were returned.",
                    false);
            }
            catch (Exception ex)
            {
                SetStatus("Failed: " + ex.Message, true);
                var telemetry = _services.GetService<ITelemetryService>();
                if (telemetry != null) telemetry.TrackError(ex, "Settings.TestConnection");
            }
            finally
            {
                SetBusy(false);
            }
        }

        private void OnSave(object sender, EventArgs e)
        {
            var url = _baseUrl.Text.Trim();
            if (_liveMode.Checked && !IsAcceptableUrl(url))
            {
                SetStatus("Enter an absolute http:// or https:// address for the API.", true);
                _baseUrl.Focus();
                return;
            }

            var typed = _token.Text.Trim();
            if (typed.Length > 0)
            {
                try
                {
                    WindowsCredentialTokenProvider.WriteCredential(
                        WindowsCredentialTokenProvider.TargetName, typed);
                }
                catch (Exception ex)
                {
                    SetStatus("Could not store the token: " + ex.Message, true);
                    return;
                }
            }

            _settings.UseDemoData = _demoMode.Checked;
            _settings.ApiBaseUrl = url;
            _settings.CacheFreshnessMinutes = (int)_freshness.Value;

            try
            {
                _settings.Save();
            }
            catch (Exception ex)
            {
                SetStatus("Could not save settings: " + ex.Message, true);
                return;
            }

            var telemetry = _services.GetService<ITelemetryService>();
            if (telemetry != null) telemetry.TrackEvent("Settings.Saved");

            DialogResult = DialogResult.OK;
            Close();
        }

        private static bool IsAcceptableUrl(string value)
        {
            Uri uri;
            return Uri.TryCreate(value, UriKind.Absolute, out uri) &&
                   (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps);
        }
    }
}
