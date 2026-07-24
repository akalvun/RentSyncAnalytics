using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Extensions.DependencyInjection;
using Office = Microsoft.Office.Core;
using RentSync.Core.Services;
using RentSync.Core.Telemetry;

namespace RentSync.AddIn
{
    /// <summary>
    /// Ribbon callbacks. This class demonstrates the single most important
    /// VSTO skill from the job ad: "multi-threading, async/await, avoiding
    /// UI blocking within Excel".
    ///
    /// The pattern in every handler:
    ///   1. Disable buttons (ribbon invalidate) so the user can't double-fire.
    ///   2. await the Core service — I/O runs off the UI thread, Excel stays
    ///      responsive (user can keep typing).
    ///   3. After await we're back on the captured UI context, so touching
    ///      the Excel COM object model is safe.
    ///   4. Errors go to telemetry + a friendly dialog; never crash Excel.
    /// </summary>
    [ComVisible(true)]
    public class RibbonController : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;
        private bool _busy;

        private static T Get<T>() => ThisAddIn.Services.GetRequiredService<T>();

        public string GetCustomUI(string ribbonId) =>
            Properties.Resources.RibbonUI; // RibbonUI.xml embedded as resource

        public void Ribbon_Load(Office.IRibbonUI ribbonUi) => _ribbon = ribbonUi;

        public bool GetActionsEnabled(Office.IRibbonControl control) => !_busy;

        // ---------------- handlers ----------------

        public async void OnRefreshClick(Office.IRibbonControl control) =>
            await RunAsync("Ribbon.Refresh", async () =>
            {
                var data = await Get<IRentRollDataService>().GetRentRollAsync(forceRefresh: false);
                ExcelActions.WriteRentRoll(data);   // back on UI thread here
            });

        public async void OnForceRefreshClick(Office.IRibbonControl control) =>
            await RunAsync("Ribbon.ForceRefresh", async () =>
            {
                var data = await Get<IRentRollDataService>().GetRentRollAsync(forceRefresh: true);
                ExcelActions.WriteRentRoll(data);
            });

        public async void OnSummaryClick(Office.IRibbonControl control) =>
            await RunAsync("Ribbon.Summary", async () =>
            {
                var summaries = await Get<IRentAnalyticsService>().GetPropertySummariesAsync();
                ExcelActions.WriteSummaries(summaries);
            });

        public async void OnExpiringClick(Office.IRibbonControl control) =>
            await RunAsync("Ribbon.Expiring", async () =>
            {
                var asOf = DateOnly.FromDateTime(DateTime.Today);
                var expiring = await Get<IRentAnalyticsService>()
                    .GetExpiringLeasesAsync(asOf, withinMonths: 6);
                ExcelActions.WriteRentRoll(expiring, sheetName: "Expiring Leases");
            });

        public void OnSettingsClick(Office.IRibbonControl control)
        {
            Get<ITelemetryService>().TrackEvent("Ribbon.SettingsOpened");
            using (var form = new SettingsForm())
                form.ShowDialog();
        }

        // ---------------- shared plumbing ----------------

        private async Task RunAsync(string operationName, Func<Task> action)
        {
            if (_busy) return;
            SetBusy(true);

            var telemetry = Get<ITelemetryService>();
            using (var op = telemetry.StartOperation(operationName))
            {
                try
                {
                    await action(); // ConfigureAwait(true) implied: resume on UI thread
                    op.Succeed();
                }
                catch (Exception ex)
                {
                    op.Fail(ex);
                    telemetry.TrackError(ex, operationName);
                    MessageBox.Show(
                        "The operation failed: " + ex.Message +
                        "\r\n\r\nDetails were written to the RentSync log.",
                        "RentSync Analytics",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                finally
                {
                    SetBusy(false);
                }
            }
        }

        private void SetBusy(bool busy)
        {
            _busy = busy;
            _ribbon?.Invalidate(); // re-queries GetActionsEnabled
        }
    }
}
