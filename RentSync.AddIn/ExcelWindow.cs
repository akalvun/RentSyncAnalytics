using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace RentSync.AddIn
{
    /// <summary>
    /// Supplies Excel's main window as an <see cref="IWin32Window"/> owner
    /// for WinForms dialogs.
    ///
    /// A modal dialog shown without an owner belongs to no window in the
    /// z-order. Clicking the workbook can put Excel in front of the dialog
    /// while the modal loop still swallows input — from the user's side,
    /// Excel has frozen for no reason. Passing the host window as owner
    /// keeps the dialog on top of Excel, centres it correctly, and lets
    /// Windows minimise the pair together.
    /// </summary>
    internal static class ExcelWindow
    {
        public static IWin32Window Owner()
        {
            // Preferred: ask Excel itself.
            try
            {
                var hwnd = new IntPtr(Globals.ThisAddIn.Application.Hwnd);
                if (hwnd != IntPtr.Zero)
                    return new Win32Window(hwnd);
            }
            catch
            {
                // Application.Hwnd throws while Excel is shutting down or in
                // some embedded hosting scenarios; fall through.
            }

            // Fallback: the process main window. Same handle in practice.
            try
            {
                var hwnd = Process.GetCurrentProcess().MainWindowHandle;
                if (hwnd != IntPtr.Zero)
                    return new Win32Window(hwnd);
            }
            catch
            {
                // Ignore — an unowned dialog is still better than no dialog.
            }

            return null;
        }

        private sealed class Win32Window : IWin32Window
        {
            private readonly IntPtr _handle;
            public Win32Window(IntPtr handle) { _handle = handle; }
            public IntPtr Handle { get { return _handle; } }
        }
    }
}
