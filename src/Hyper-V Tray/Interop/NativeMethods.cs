using System;
using System.Runtime.InteropServices;

namespace HyperVTray.Interop
{
    internal static class NativeMethods
    {
        #region Methods

        #region Private Static

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hwnd);
        [DllImport("user32.dll")]
        private static extern int GetSystemMetricsForDpi(int smIndex, uint dpi);

        #endregion

        #region Internal Static

        internal static int GetTrayIconWidth(IntPtr handle)
        {
            var dpi = GetDpiForWindow(handle);
            return GetSystemMetricsForDpi(49, dpi); // 49 == SM_CXSMICON
        }

        #endregion

        #endregion
    }
}