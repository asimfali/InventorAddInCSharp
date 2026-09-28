using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;

namespace InvDoc
{
    class unmanaged
    {
        [DllImport("shell32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SHGetPathFromIDListW(IntPtr pidl, [MarshalAs(UnmanagedType.LPTStr)] StringBuilder pszPath);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        public static string GetPathFromPIDL(byte[] byteCode)
        {
            //MAX_PATH = 260 
            StringBuilder builder = new StringBuilder(260);

            IntPtr ptr = IntPtr.Zero;
            GCHandle h0 = GCHandle.Alloc(byteCode, GCHandleType.Pinned);
            try
            {
                ptr = h0.AddrOfPinnedObject();
            }
            finally
            {
                h0.Free();
            }

            SHGetPathFromIDListW(ptr, builder);

            return builder.ToString();
        }

        public static void setForeground(string name)
        {
            Process[] p = Process.GetProcessesByName(name);
            if (p.Length > 0)
            {
                SetForegroundWindow(p[0].MainWindowHandle);
            }
        }
    }
}
