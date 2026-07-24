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
            collection.AddRentSyncCore();

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
            try
            {
                var telemetry = _services?.GetService<ITelemetryService>();
                telemetry?.TrackEvent("AddIn.Shutdown");
                // Bounded flush: never hang Excel's exit.
                telemetry?.FlushAsync(new CancellationTokenSource(
                    TimeSpan.FromSeconds(2)).Token).GetAwaiter().GetResult();
            }
            catch { /* shutdown must be silent */ }
            finally
            {
                _services?.Dispose();
            }
        }

        /// <summary>Expose the VBA interop bridge: in VBA use
        /// Application.COMAddIns("RentSync.AddIn").Object</summary>
        protected override object RequestComAddInAutomationService()
        {
            if (_vbaBridge == null) _vbaBridge = new VbaBridge();
            return _vbaBridge;
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
