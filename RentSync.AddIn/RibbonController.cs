using System;
using System.Drawing;
using System.IO;
using System.Reflection;
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
    ///   2. await the Core service - I/O runs off the UI thread, Excel stays
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

        public string GetCustomUI(string ribbonId) => LoadRibbonXml();

        /// <summary>
        /// Read RibbonUI.xml straight out of the assembly's embedded resources.
        /// Avoids depending on a generated Properties.Resources entry, which is
        /// fragile to set up by hand in a classic project. The resource name is
        /// "{RootNamespace}.RibbonUI.xml" = "RentSync.AddIn.RibbonUI.xml".
        /// </summary>
        private static string LoadRibbonXml()
        {
            var asm = Assembly.GetExecutingAssembly();
            const string resourceName = "RentSync.AddIn.RibbonUI.xml";

            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    // Diagnostic aid: if the name is wrong, show what IS embedded.
                    var available = string.Join(", ", asm.GetManifestResourceNames());
                    throw new InvalidOperationException(
                        "Embedded ribbon resource '" + resourceName +
                        "' not found. Available resources: " + available);
                }
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUi) => _ribbon = ribbonUi;

        public bool GetActionsEnabled(Office.IRibbonControl control) => !_busy;

        /// <summary>
        /// loadImage callback (referenced by loadImage="OnLoadImage" in the XML).
        /// Excel passes the image id from each control's image="..." attribute and
        /// expects a Bitmap back. We serve PNGs embedded in the assembly, so the
        /// custom RentSync logo can sit in the ribbon like a brand mark.
        /// </summary>
        public System.Drawing.Bitmap OnLoadImage(string imageId)
        {
            var asm = Assembly.GetExecutingAssembly();

            // Try the expected name first.
            var expected = "RentSync.AddIn.Assets." + imageId + ".png";
            var stream = asm.GetManifestResourceStream(expected);

            // Fallback: match by suffix, so a different root namespace still works
            // (e.g. "Assets.logo32.png" embedded under some other prefix).
            if (stream == null)
            {
                foreach (var name in asm.GetManifestResourceNames())
                {
                    if (name.EndsWith(imageId + ".png", StringComparison.OrdinalIgnoreCase))
                    {
                        stream = asm.GetManifestResourceStream(name);
                        break;
                    }
                }
            }

            if (stream == null) return null;
            using (stream)
                return new System.Drawing.Bitmap(stream);
        }

        // ---------------- handlers ----------------

        public async void OnRefreshClick(Office.IRibbonControl control) =>
            await RunAsync("Ribbon.Refresh", async () =>
            {
                var data = await Get<IRentRollDataService>().GetRentRollAsync(forceRefresh: false);
                ExcelActions.WriteRentRoll(data);
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
                using (var form = new SettingsForm(ThisAddIn.Services))
                form.ShowDialog(ExcelWindow.Owner());
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
                    await action(); // resumes on UI thread
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
            _ribbon?.Invalidate();
        }
    }
}
