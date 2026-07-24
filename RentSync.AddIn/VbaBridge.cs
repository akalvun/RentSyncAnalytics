using System;
using System.Runtime.InteropServices;
using Microsoft.Extensions.DependencyInjection;
using RentSync.Core.Services;
using RentSync.Core.Telemetry;

namespace RentSync.AddIn
{
    /// <summary>
    /// ComVisible bridge so existing VBA macros can drive the add-in —
    /// the "VBA and interoperability with VSTO" job requirement.
    ///
    /// From VBA:
    ///   Dim rs As Object
    ///   Set rs = Application.COMAddIns("RentSync.AddIn").Object
    ///   rs.RefreshData            ' cache-first load into the Rent Roll sheet
    ///   MsgBox rs.GetVersion
    /// </summary>
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IVbaBridge
    {
        string GetVersion();
        void RefreshData();
        void ForceRefresh();
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class VbaBridge : IVbaBridge
    {
        public string GetVersion() =>
            typeof(VbaBridge).Assembly.GetName().Version.ToString();

        public void RefreshData() => Run(forceRefresh: false);

        public void ForceRefresh() => Run(forceRefresh: true);

        private static void Run(bool forceRefresh)
        {
            var telemetry = ThisAddIn.Services.GetRequiredService<ITelemetryService>();
            telemetry.TrackEvent("Vba.Refresh",
                new System.Collections.Generic.Dictionary<string, object>
                    { ["forceRefresh"] = forceRefresh });

            // VBA calls arrive on the UI thread; keep it responsive by doing
            // the I/O on the pool and marshalling the Excel write back.
            System.Threading.Tasks.Task.Run(async () =>
            {
                try
                {
                    var data = await ThisAddIn.Services
                        .GetRequiredService<IRentRollDataService>()
                        .GetRentRollAsync(forceRefresh)
                        .ConfigureAwait(false);

                    ThisAddIn.UiContext.Post(_ => ExcelActions.WriteRentRoll(data), null);
                }
                catch (Exception ex)
                {
                    telemetry.TrackError(ex, "Vba.Refresh");
                }
            });
        }
    }
}
