using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PartsCounter
{
   public static class SingleInstance
    {
        public static readonly int WM_SHOWFIRSTINSTANCE =Win32API.RegisterWindowMessage("WM_SHOWFIRSTINSTANCE|{0}",ProgramInfo.AssemblyGuid);
        static Mutex mutex;
        static public bool Start()
        {
            bool onlyInstance = false;
            string mutexName = String.Format("Local\\{0}", ProgramInfo.AssemblyGuid);
            mutex = new Mutex(true, mutexName, out onlyInstance);
            return onlyInstance;
        }
        static public void ShowFirstInstance()
        {
            Win32API.PostMessage((IntPtr)Win32API.HWND_BROADCAST, WM_SHOWFIRSTINSTANCE, IntPtr.Zero, IntPtr.Zero);
        }
        static public void Stop()
        {
            mutex.ReleaseMutex();
        }
    }
}
