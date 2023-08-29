using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace AutomationLib
{
    public class Class1
    {
        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPhysicalPoint(int x, int y);

        public static IntPtr GetWindowHandle(int x, int y)
        {
          
            IntPtr i = WindowFromPhysicalPoint(x, y);
            return i;
        }
    }
}
