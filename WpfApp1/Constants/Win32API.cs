using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace WpfApp1.Constants
{
    public class Win32API
    {
        [DllImport("user32.dll")]
        public extern static IntPtr GetDC(IntPtr hWnd);
        [DllImport("user32.dll", EntryPoint = "WindowFromPoint")]
        public static extern IntPtr WindowFromPoint(int x, int y);

        [DllImport("user32")]
        public static extern bool InvalidateRect(IntPtr hwnd, IntPtr rect, bool bErase);

        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out System.Drawing.Point point);
    }
}
