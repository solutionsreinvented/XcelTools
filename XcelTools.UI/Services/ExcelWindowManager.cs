using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace XcelTools.Licensing.Services
{
    public class ExcelWindowManager
    {
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        private const int SW_HIDE = 0;
        private const int SW_SHOW = 5;

        public static IntPtr GetExcelMainWindowHandle()
        {
            return Process.GetCurrentProcess().MainWindowHandle;
        }

        public static void HideExcelWindow()
        {
            var handle = GetExcelMainWindowHandle();
            if (handle != IntPtr.Zero && IsWindowVisible(handle))
            {
                ShowWindow(handle, SW_HIDE);
            }
        }

        public static void ShowExcelWindow()
        {
            var handle = GetExcelMainWindowHandle();
            if (handle != IntPtr.Zero)
            {
                ShowWindow(handle, SW_SHOW);
            }
        }
    }
}
