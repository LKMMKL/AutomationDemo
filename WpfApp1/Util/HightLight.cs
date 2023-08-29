using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using UIAutomationClient;
using WpfApp1.Constants;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

namespace WpfApp1.Util
{
    public class HightLight
    {
        public static Task draw_task;
        public static CancellationTokenSource tokenSource = new CancellationTokenSource();
        public static Action action;
        public static void DrawHightLight(tagRECT r)
        {
            if (draw_task != null )
            {
                tokenSource.Cancel();
                ResetToken();
            }
            var ct = tokenSource.Token;
            Rectangle rect = new Rectangle(r.left, r.top, r.right-r.left+2, r.bottom-r.top+2);
            draw_task = new Task(() =>
            {
                IntPtr desktop = Win32API.GetDC(IntPtr.Zero);
                while (!ct.IsCancellationRequested)
                {
                    using (Graphics g = Graphics.FromHdc(desktop))
                    {
                        Pen myPen = new Pen(System.Drawing.Color.Red, 5);
                        g.DrawRectangle(myPen, rect);
                        g.Dispose();
                    }
                }
                var r1 = IntPtr.Zero;
                IntPtr pnt = Marshal.AllocHGlobal(Marshal.SizeOf(rect));
                Marshal.StructureToPtr(r1, pnt, false);
                //InvalidateRect(IntPtr.Zero,r , true);
                Win32API.InvalidateRect(IntPtr.Zero, r1, true);
                
            });
            
            draw_task.Start();
        }

        public static void ResetToken()
        {
            tokenSource = new CancellationTokenSource();
        }

        static HightLight()
        {
            System.Windows.Forms.Timer time = new System.Windows.Forms.Timer();
            time.Tick += Time_Tick;
            time.Interval = 3000;
            time.Start();
        }

        private static void Time_Tick(object sender, EventArgs e)
        {
            Point point = new Point();
            tagPOINT tp = new tagPOINT();
            Win32API.GetCursorPos(out point);
            tp.y = point.Y;
            tp.x = point.X;
            // 指针在主窗口内
            if (false) {
            }
            else
            {
                //List<EleInfo> eles = UIControlAssist.allEle;
                //if(eles.Count > 0)
                //{
                //    for(int i=0;i<eles.Count;i++)
                //    {
                //        tagRECT? rect = eles[i].ContainPoint(point);
                //        if (rect != null) {
                //            DrawHightLight((tagRECT)rect);
                //        }
                //    }
                //}

                CUIAutomation uia = new CUIAutomation();
                IUIAutomationElement ele = uia.ElementFromPoint(tp);
                if (ele != null)
                {
                    DrawHightLight(ele.CurrentBoundingRectangle);
                }
            }
        }
    }
}
