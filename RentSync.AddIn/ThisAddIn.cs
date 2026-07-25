using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using RentSync.Core;
using RentSync.Core.Services;
using RentSync.Core.Telemetry;

namespace RentSync.AddIn
{
    /// <summary>
    /// VSTO entry point. Responsibilities kept deliberately thin:
    ///  1. Build the DI container from RentSync.Core (single composition root).
    ///  2. Capture the WinForms SynchronizationContext so async work can
    ///     marshal results back to Excel's UI thread (COM objects must only
    ///     be touched from that thread).
    ///  3. Flush telemetry on shutdown.
    ///  4. Expose a ComVisible object to VBA (RequestComAddInAutomationService).
    /// </summary>
    public partial class ThisAddIn
    {
        private ServiceProvider _services;
        private VbaBridge _vbaBridge;

        /// <summary>Excel's UI thread context; set once in Startup.</summary>
        public static SynchronizationContext UiContext { get; private set; }

        public static IServiceProvider Services { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // VSTO guarantees Startup runs on the UI thread.
            UiContext = SynchronizationContext.Current
                        ?? new WindowsFormsSynchronizationContext();

            var appData = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "RentSync");

            var collection = new ServiceCollection();
            collection.AddRentSyncLogging(Path.Combine(appData, "logs"));

            // Point Core at the bundled sample JSON (demo mode). VSTO shadow-copies
            // the add-in into an assembly cache, so Assembly.Location points at that
            // cache - not where our Data folder is. CodeBase preserves the original
            // deployment location; strip the file:// URI prefix to get a real path.
            var codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            var uri = new Uri(codeBase);
            var addinDir = Path.GetDirectoryName(uri.LocalPath);
            var samplePath = Path.Combine(addinDir, "Data", "rentroll.sample.json");

            collection.AddRentSyncCore(configureApi: api =>
            {
                api.UseLocalSample = true;
                api.LocalSamplePath = samplePath;
            });

            // Swap the anonymous token provider for the Windows Credential
            // Manager implementation (job requirement: auth & security).
            collection.Replace(ServiceDescriptor.Singleton<IAuthTokenProvider,
                WindowsCredentialTokenProvider>());

            _services = collection.BuildServiceProvider();
            Services = _services;

            // Warm up the cache off the UI thread; errors are logged, never thrown
            // into Excel's startup path.
            var cache = _services.GetRequiredService<ICacheService>();
            var telemetry = _services.GetRequiredService<ITelemetryService>();
            System.Threading.Tasks.Task.Run(async () =>
            {
                try { await cache.InitializeAsync().ConfigureAwait(false); }
                catch (Exception ex) { telemetry.TrackError(ex, "AddIn.CacheInit"); }
            });

            telemetry.TrackEvent("AddIn.Started");
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            var provider = _services;
            _services = null;
            if (provider == null) return;

            try
            {
                (provider.GetService<ITelemetryService>())?.TrackEvent("AddIn.Shutdown");
            }
            catch { /* ignore */ }

            // Dispose off the UI thread and DO NOT block on it. Blocking here
            // (.Wait / .GetResult) would deadlock: the async disposal of the
            // telemetry service tries to resume on the UI SynchronizationContext
            // captured at startup, but that thread is stuck waiting on us.
            // Excel's exit must never wait for telemetry to flush.
            System.Threading.Tasks.Task.Run(async () =>
            {
                try { await provider.DisposeAsync().ConfigureAwait(false); }
                catch { /* shutdown must be silent */ }
            });
        }

        /// <summary>Expose the VBA interop bridge: in VBA use
        /// Application.COMAddIns("RentSync.AddIn").Object</summary>
        protected override object RequestComAddInAutomationService()
        {
            if (_vbaBridge == null) _vbaBridge = new VbaBridge();
            return _vbaBridge;
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility
            CreateRibbonExtensibilityObject()
        {
            return new RibbonController();
        }

        #region VSTO generated code
        // The generated ThisAddIn.Designer.cs calls this.InternalStartup() from
        // FinishInitialization(), but does NOT define it - this template expects
        // the method to live here. Wire the Startup/Shutdown events.
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
