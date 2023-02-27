using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace PartsCounter
{
    static public class Win32API
    {
        [DllImport("user32")]
        public static extern int RegisterWindowMessage(string message);
        public static int RegisterWindowMessage(string format,params object[] args)
        {
            string message = string.Format(format, args);
            return RegisterWindowMessage(message);
        }
        public const int HWND_BROADCAST = 0xffff;
        public const int SW_SHOWNORMAL = 1;
        public const int SW_RESTORE = 9;
        [DllImport("user32")]
        public static extern bool PostMessage(IntPtr hwnd, int msg, IntPtr wparam, IntPtr lparam);
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        public static void ShowToFront(IntPtr window)
        {
            ShowWindow(window, SW_RESTORE);
            SetForegroundWindow(window);
        }

    }
}
