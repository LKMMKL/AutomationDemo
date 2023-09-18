using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices; //调用WINDOWS API函数时要用到
using Microsoft.Win32; //写入注册表时要用到
using System.Windows.Input;
using System.Windows.Forms;
using KeyEventHandler = System.Windows.Forms.KeyEventHandler;
using KeyEventArgs = System.Windows.Forms.KeyEventArgs;
using MouseEventArgs = System.Windows.Forms.MouseEventArgs;
using MouseEventHandler = System.Windows.Forms.MouseEventHandler;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Threading;
using System.Windows;
using System.Windows.Media;
using Point = System.Drawing.Point;
using UIAutomationClient;

namespace WpfApp1
{
    public partial class KeyboardHook
    {
        public string labels { get; set; }
        public KeyboardHook()
        {
            labels = "KeyboardHook";
        }
        public delegate int HookProc(int nCode, Int32 wParam, IntPtr lParam);
        HookProc KeyboardHookProcedure; //声明KeyboardHookProcedure作为HookProc类型
        static int hKeyboardHook = 0; //声明键盘钩子处理的初始值
        //值在Microsoft SDK的Winuser.h里查询
        // http://www.bianceng.cn/Programming/csharp/201410/45484.htm
        //定义为鼠标钩子
        public int WH_MOUSE_LL = 14;
        
        //键盘结构
        [StructLayout(LayoutKind.Sequential)]
        public class KeyboardHookStruct
        {
            public int vkCode;  //定一个虚拟键码。该代码必须有一个价值的范围1至254
            public int scanCode; // 指定的硬件扫描码的关键
            public int flags;  // 键标志
            public int time; // 指定的时间戳记的这个讯息
            public int dwExtraInfo; // 指定额外信息相关的信息
        }
        //使用此功能，安装了一个钩子
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern int SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hInstance, int threadId);


        //调用此函数卸载钩子
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern bool UnhookWindowsHookEx(int idHook);


        //使用此功能，通过信息钩子继续下一个钩子
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern int CallNextHookEx(int idHook, int nCode, Int32 wParam, IntPtr lParam);

        // 取得当前线程编号（线程钩子需要用到）
        [DllImport("kernel32.dll")]
        static extern int GetCurrentThreadId();

        //使用WINDOWS API函数代替获取当前实例的函数,防止钩子失效
        [DllImport("kernel32.dll")]
        public static extern IntPtr GetModuleHandle(string name);
        private int hookID = 0;

        MainWindow w;
        public void Start(MainWindow window)
        {
            Timer time = new Timer();
            
            time.Tick += Time_Tick;
            time.Interval = 3000;
            time.Start();
            w = window;
            //// 安装键盘钩子
            //if (hKeyboardHook == 0)
            //{

            //    KeyboardHookProcedure = new HookProc(KeyboardHookProc);
            //    IntPtr handle = System.Runtime.InteropServices.Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0]);
            //    hKeyboardHook = SetWindowsHookEx(WH_MOUSE_LL, KeyboardHookProcedure, handle, 0);
            //    GC.KeepAlive(KeyboardHookProcedure);
            //    //hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, KeyboardHookProcedure, Marshal.GetHINSTANCE(Assembly.GetExecutingAssembly().GetModules()[0]), 0);
            //    //************************************
            //    //键盘线程钩子
            //    //SetWindowsHookEx( 2,KeyboardHookProcedure, IntPtr.Zero, GetCurrentThreadId());//指定要监听的线程idGetCurrentThreadId(),
            //    //键盘全局钩子,需要引用空间(using System.Reflection;)
            //    //SetWindowsHookEx( 13,MouseHookProcedure,Marshal.GetHINSTANCE(Assembly.GetExecutingAssembly().GetModules()[0]),0);
            //    //
            //    //关于SetWindowsHookEx (int idHook, HookProc lpfn, IntPtr hInstance, int threadId)函数将钩子加入到钩子链表中，说明一下四个参数：
            //    //idHook 钩子类型，即确定钩子监听何种消息，上面的代码中设为2，即监听键盘消息并且是线程钩子，如果是全局钩子监听键盘消息应设为13，
            //    //线程钩子监听鼠标消息设为7，全局钩子监听鼠标消息设为14。lpfn 钩子子程的地址指针。如果dwThreadId参数为0 或是一个由别的进程创建的
            //    //线程的标识，lpfn必须指向DLL中的钩子子程。 除此以外，lpfn可以指向当前进程的一段钩子子程代码。钩子函数的入口地址，当钩子钩到任何
            //    //消息后便调用这个函数。hInstance应用程序实例的句柄。标识包含lpfn所指的子程的DLL。如果threadId 标识当前进程创建的一个线程，而且子
            //    //程代码位于当前进程，hInstance必须为NULL。可以很简单的设定其为本应用程序的实例句柄。threaded 与安装的钩子子程相关联的线程的标识符
            //    //如果为0，钩子子程与所有的线程关联，即为全局钩子
            //    //************************************
            //    //如果SetWindowsHookEx失败
            //    if (hKeyboardHook == 0)
            //    {
            //        Stop();
            //        throw new Exception("安装键盘钩子失败");
            //    }
            //}
        }
        public static List<IUIAutomationElement> list;
        private void Time_Tick(object sender, EventArgs e)
        {
            Point p = new Point();
            GetPhysicalCursorPos(out p);
            DataTemplate t = new DataTemplate();
            Task.Run(() =>
            {
                //list = UIControlAssist.GetDesktop();
                //w.Dispatcher.Invoke(() =>
                //{
                //    w.tree.ItemsSource = list;
                    
                //});
               

            });
            Console.WriteLine("=============");
        }

        public void Stop()
        {
            bool retKeyboard = true;


            if (hKeyboardHook != 0)
            {
                retKeyboard = UnhookWindowsHookEx(hKeyboardHook);
                hKeyboardHook = 0;
            }

            if (!(retKeyboard)) throw new Exception("卸载钩子失败！");
        }
        //ToAscii职能的转换指定的虚拟键码和键盘状态的相应字符或字符
        [DllImport("user32")]
        public static extern int ToAscii(int uVirtKey, //[in] 指定虚拟关键代码进行翻译。
                                         int uScanCode, // [in] 指定的硬件扫描码的关键须翻译成英文。高阶位的这个值设定的关键，如果是（不压）
                                         byte[] lpbKeyState, // [in] 指针，以256字节数组，包含当前键盘的状态。每个元素（字节）的数组包含状态的一个关键。如果高阶位的字节是一套，关键是下跌（按下）。在低比特，如果设置表明，关键是对切换。在此功能，只有肘位的CAPS LOCK键是相关的。在切换状态的NUM个锁和滚动锁定键被忽略。
                                         byte[] lpwTransKey, // [out] 指针的缓冲区收到翻译字符或字符。
                                         int fuState); // [in] Specifies whether a menu is active. This parameter must be 1 if a menu is active, or 0 otherwise.


        [StructLayout(LayoutKind.Sequential)]
        public class MouseHookStruct
        {
            public POINT pt; // 鼠标位置
            public int hWnd;
            public int wHitTestCode;
            public int dwExtraInfo;
        }

        /// <summary>
        /// 鼠标位置结构
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public class POINT
        {
            public int x;
            public int y;
        }

        //相关鼠标事件
        public event MouseEventHandler MouseDown;
        public event MouseEventHandler MouseUp;

        //相关动作
        public const int WM_MOUSEMOVE = 0x200; // 鼠标移动
        public const int WM_LBUTTONDOWN = 0x201;// 鼠标左键按下
        public const int WM_RBUTTONDOWN = 0x204;// 鼠标右键按下
        public const int WM_MBUTTONDOWN = 0x207;// 鼠标中键按下
        public const int WM_LBUTTONUP = 0x202;// 鼠标左键抬起
        public const int WM_RBUTTONUP = 0x205;// 鼠标右键抬起
        public const int WM_MBUTTONUP = 0x208;// 鼠标中键抬起

        //hookid

        private int KeyboardHookProc(int nCode, Int32 wParam, IntPtr lParam)
        {
            return 1;
            //// 侦听键盘事件
            //if ((nCode >= 0))
            //{
            //    MouseHookStruct hookStruct = (MouseHookStruct)Marshal.PtrToStructure(lParam, typeof(MouseHookStruct));

            //    MouseEventArgs e = null;
            //    switch (wParam)
            //    {
            //        case WM_LBUTTONDOWN:
            //            e = new MouseEventArgs(MouseButtons.Left, 1, hookStruct.pt.x, hookStruct.pt.y, 0);
            //            //MouseMove(this, e);
            //            break;
            //        case WM_RBUTTONDOWN:
            //            e = new MouseEventArgs(MouseButtons.Right, 1, hookStruct.pt.x, hookStruct.pt.y, 0);
            //            Stop();
            //            break;
            //        case WM_LBUTTONUP:
            //            e = new MouseEventArgs(MouseButtons.Left, 1, hookStruct.pt.x, hookStruct.pt.y, 0);
            //            //MouseUp(this, e);
            //            break;
            //        case WM_RBUTTONUP:
            //            e = new MouseEventArgs(MouseButtons.Right, 1, hookStruct.pt.x, hookStruct.pt.y, 0);
            //            MouseUp(this, e);
            //            break;
            //        case WM_MOUSEMOVE:
            //            e = new MouseEventArgs(MouseButtons.Right, 1, hookStruct.pt.x, hookStruct.pt.y, 0);
            //            MouseMove(this,e);
            //            break;
            //        default:
            //            break;
            //    }
            //}
            ////如果返回1，则结束消息，这个消息到此为止，不再传递。
            ////如果返回0或调用CallNextHookEx函数则消息出了这个钩子继续往下传递，也就是传给消息真正的接受者
            //return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
        }
        Graphics g = Graphics.FromHwnd(IntPtr.Zero);
        [DllImport("user32")]
        private static extern bool InvalidateRect(IntPtr hwnd, IntPtr rect, bool bErase);

        [DllImport("user32.dll")]
        public static extern bool GetPhysicalCursorPos(out System.Drawing.Point point);
        [DllImport("user32.dll")]
        static extern bool GetCursorPos(out System.Drawing.Point point);
        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPhysicalPoint(int x, int y);
        [DllImport("user32.dll")]
        public static extern Boolean ClientToScreen(IntPtr hwnd, ref System.Drawing.Point point);
        private void MouseMove(object sender, MouseEventArgs e)
        {
            //w.Height = 50;
            //w.Width = 100;
            //w.Left = e.X + 10;
            //w.Top = e.Y + 10;
            Task.Run(() =>
            {
                UIControlAssist.GetPointTarget(e.X, e.Y);
                //System.Drawing.Point p = new System.Drawing.Point();
                //GetPhysicalCursorPos(out p);
                //IntPtr hwnd = WindowFromPhysicalPoint(p.X, p.Y);
                //w.Dispatcher.Invoke(()=>
                //{
                //    w.listbox.Items.Add($"{p.X} {p.Y} {hwnd}");
                //    w.listbox.ScrollIntoView(w.listbox.Items.GetItemAt(w.listbox.Items.Count - 1));
                //});
                
            });
            //IntPtr r = IntPtr.Zero;
            //Rectangle re = new Rectangle(e.X, e.Y, 600, 400);
            //IntPtr pnt = Marshal.AllocHGlobal(Marshal.SizeOf(re));
            //Marshal.StructureToPtr(r, pnt, false);
            //InvalidateRect(IntPtr.Zero,r , true);
            //g.DrawRectangle(new Pen(Color.Red), new Rectangle(e.X, e.Y, 600, 400));
            //var str = string.Format("在{0},{1}位置按下了鼠标{2}键", e.X, e.Y, e.Button.ToString());

        }
        public static IntPtr h;
        [DllImport("Kernel32.dll")]
        public static extern int GetLastError();
        [DllImport("user32", EntryPoint = "GetDesktopWindow")]
        public static extern IntPtr GetDesktopWindow();

        [DllImport("user32", EntryPoint = "GetDC")]
        public static extern IntPtr GetDC(IntPtr h);
        public static void DrawHighLight(IntPtr hwnd, tagRECT rect)
        {
            Rectangle re = new Rectangle(rect.left, rect.top, rect.right, rect.bottom);
            IntPtr pnt = Marshal.AllocHGlobal(Marshal.SizeOf(re));
            if(h != null)InvalidateRect(IntPtr.Zero, pnt, true);
            IntPtr hdc = GetDC(IntPtr.Zero);
            h = hdc;
            Graphics g = Graphics.FromHdc(hdc);
            g.DrawRectangle(new System.Drawing.Pen(System.Drawing.Color.Red), re.Left, re.Top, re.Right-re.Left, re.Bottom-re.Top);
        }

        ~KeyboardHook()
        {
            Stop();
        }
    }
}
