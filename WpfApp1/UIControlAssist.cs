using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using UIAutomationClient;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using Point = System.Drawing.Point;

namespace WpfApp1
{
    public class EleInfo : INotifyPropertyChanged
    {
        public string Name { get; set; }
        public string ClassName { get; set; }
        public string AutomationId { get; set; }
        public tagRECT rect { get; set; }
        public string RuntimeId { get; set; }
        public string rootId;
        public int level { get; set; }
        public List<EleInfo> childs { get; set; }
        public bool check;
        public bool checkValue {
            get
            {
                return check;
            }
            set {
                check = !check;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(showRefresh)));
            }
        }
        public Visibility showRefresh
        {
            get
            {
                return checkValue ? Visibility.Visible : Visibility.Collapsed;
            }
            set
            {

                showCheckbox = value;
            }
        }
        public Visibility showCheckbox { get
            {
                return level == 1 ? Visibility.Visible : Visibility.Collapsed;
            }
            set {
                showCheckbox = value;
            }
        }
        public IUIAutomationElement curr; //当前节点


        public event PropertyChangedEventHandler PropertyChanged;

        public EleInfo()
        {
            
        }
        public EleInfo(IUIAutomationElement root, int level = 0)
        {
            Name = root.CurrentName;
            ClassName = root.CurrentClassName;
            rect = root.CurrentBoundingRectangle;
            AutomationId = root.CurrentAutomationId;
            RuntimeId = UIControlAssist.GetRuntimeIdStr(root.GetRuntimeId());
            rootId = UIControlAssist.GetRuntimeIdStr(root.GetRuntimeId());
            curr = root;
            this.level = ++level;
        }
        public EleInfo(IUIAutomationElement root,IUIAutomationElement ele, int level=0)
         {
            try
            {   
                Name = ele.CurrentName == string.Empty ? "\"\"" : ele.CurrentName;
                ClassName = ele.CurrentClassName;
                rect = ele.CurrentBoundingRectangle;
                AutomationId = ele.CurrentAutomationId;
                RuntimeId = UIControlAssist.GetRuntimeIdStr(ele.GetRuntimeId());
                rootId = UIControlAssist.GetRuntimeIdStr(root.GetRuntimeId());
                curr = ele;
                this.level = ++level;
                if (this.level < 4)
                {
                    childs = UIControlAssist.GetAllElementEx(root, curr, this.level);
                }
                else
                {
                    // 空元素
                    childs = new List<EleInfo>() { new EleInfo() };
                }
                //TimeLimitedTaskScheduler.factory.StartNew(() =>
                //{
                //    childs = UIControlAssist.GetAllElementEx(curr, this.level);
                //});
                //Task task = new Task(() =>
                //{
                //    childs = UIControlAssist.GetAllElementEx(curr, this.level);
                //});
                //UIControlAssist.TaskList.Add(task);
                //task.Start();
                //Task.Run(() =>
                //{
                //childs = UIControlAssist.GetAllElementEx(root, curr, this.level);
                //});
                UIControlAssist.map.TryAdd(RuntimeId, this);
            }
            catch(COMException ce) {
            }
            

        }

        public void Retrive()
        {
            //childs = UIControlAssist.GetAllElementEx(curr, this.level);
        }

        public tagRECT? ContainPoint(Point point)
        {
            tagRECT? tmp = null;
            if ((point.X > rect.left) && (point.X<rect.right) && (point.Y>rect.top) && (point.Y < rect.bottom))
            {
                for(int i = 0; i < childs.Count; i++)
                {
                    tmp = childs[i].ContainPoint(point);
                    if (tmp != null) break;
                }
                return tmp == null?rect:tmp;
            }
            return tmp;
        }
    }
    public class UIControlAssist
    {
        #region Enum or struct
        public enum ControlScope
        {
            Central,
            Left,
            Right,
            Top,
            Bottom
        }

        public enum ControlState
        {
            STATE_SYSTEM_UNAVAILABLE = 0x00000001,  // Disabled
            STATE_SYSTEM_SELECTED = 0x00000002,
            STATE_SYSTEM_FOCUSED = 0x00000004,
            STATE_SYSTEM_PRESSED = 0x00000008,
            STATE_SYSTEM_CHECKED = 0x00000010,
            STATE_SYSTEM_MIXED = 0x00000020,  // 3-state checkbox or toolbar button
            STATE_SYSTEM_INDETERMINATE = STATE_SYSTEM_MIXED,
            STATE_SYSTEM_READONLY = 0x00000040,
            STATE_SYSTEM_HOTTRACKED = 0x00000080,
            STATE_SYSTEM_DEFAULT = 0x00000100,
            STATE_SYSTEM_EXPANDED = 0x00000200,
            STATE_SYSTEM_COLLAPSED = 0x00000400,
            STATE_SYSTEM_BUSY = 0x00000800,
            STATE_SYSTEM_FLOATING = 0x00001000,  // Children "owned" not "contained" by parent
            STATE_SYSTEM_MARQUEED = 0x00002000,
            STATE_SYSTEM_ANIMATED = 0x00004000,
            STATE_SYSTEM_INVISIBLE = 0x00008000,
            STATE_SYSTEM_OFFSCREEN = 0x00010000,
            STATE_SYSTEM_SIZEABLE = 0x00020000,
            STATE_SYSTEM_MOVEABLE = 0x00040000,
            STATE_SYSTEM_SELFVOICING = 0x00080000,
            STATE_SYSTEM_FOCUSABLE = 0x00100000,
            STATE_SYSTEM_SELECTABLE = 0x00200000,
            STATE_SYSTEM_LINKED = 0x00400000,
            STATE_SYSTEM_TRAVERSED = 0x00800000,
            STATE_SYSTEM_MULTISELECTABLE = 0x01000000, // Supports multiple selection
            STATE_SYSTEM_EXTSELECTABLE = 0x02000000,  // Supports extended selection
            STATE_SYSTEM_ALERT_LOW = 0x04000000,  // This information is of low priority
            STATE_SYSTEM_ALERT_MEDIUM = 0x08000000,  // This information is of medium priority
            STATE_SYSTEM_ALERT_HIGH = 0x10000000,  // This information is of high priority
            STATE_SYSTEM_PROTECTED = 0x20000000,  // access to this is restricted
            STATE_SYSTEM_VALID = 0x3FFFFFFF,
        }
        #endregion

        #region Public
        [DllImport("user32")]
        private static extern bool InvalidateRect(IntPtr hwnd, IntPtr rect, bool bErase);
        private static IUIAutomationElement target;
         public static List<Task> TaskList = new List<Task>();
        public static ConcurrentDictionary<string, EleInfo> map = new ConcurrentDictionary<string, EleInfo>();

        [STAThread]
        public static void GetPointTarget(int left, int top)
        {
            CUIAutomation uia = new CUIAutomation();
            
            tagPOINT po = new tagPOINT();
            po.x = left;
            po.y = top;
            IUIAutomationElement ele = uia.ElementFromPoint(po);
            
            Task.Run(() =>
            {
                if (target != null) InvalidateRect(target.CurrentNativeWindowHandle, IntPtr.Zero, true);
                Pen pen = new Pen(Color.Red);
                tagRECT rect = ele.CurrentBoundingRectangle;
                Graphics g = Graphics.FromHwnd(ele.CurrentNativeWindowHandle);
                g.DrawRectangle(pen, rect.left, rect.top, rect.right, rect.bottom);
                target = ele;
            });
            
            Console.WriteLine(ele.CurrentName);
            Console.WriteLine($"{ele.CurrentName} {ele.CurrentBoundingRectangle.left} {ele.CurrentBoundingRectangle.top} {ele.CurrentBoundingRectangle.right} {ele.CurrentBoundingRectangle.bottom}");
        }

        public static List<EleInfo> allEle = new List<EleInfo>();
        [STAThread]
        public static List<EleInfo> GetAllElement()
        {
            List<EleInfo> eles = new List<EleInfo>();
            CUIAutomation uia = new CUIAutomation();
            IUIAutomationPropertyCondition find1_condition =
                   uia.CreatePropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId,
                    UIA_ControlTypeIds.UIA_WindowControlTypeId) as IUIAutomationPropertyCondition;
            IUIAutomationElementArray arry = uia.GetRootElement().FindAll(TreeScope.TreeScope_Children, find1_condition);
            //CountdownEvent countdownEvent = new CountdownEvent(arry.Length);
            //eles.Add(new EleInfo(uia.GetRootElement(), uia.GetRootElement(), -1));
            for (int i = 0; i < arry.Length; i++)
            {
                eles.Add(new EleInfo(arry.GetElement(i), arry.GetElement(i)));
            }
            EleInfo root = new EleInfo(uia.GetRootElement(), -1);
            root.childs = eles;
            //eles.Add(new EleInfo(arry.GetElement(1), arry.GetElement(1)));
            //eles.Add(new EleInfo(arry.GetElement(2), arry.GetElement(2)));
            //countdownEvent.Wait();
            //Task.WaitAll(TaskList.ToArray());
            //allEle = eles;
            return new List<EleInfo> { root};
        }
        public static void Refresh(EleInfo eles)
        {

            foreach (var item in map.Where(kvp => kvp.Value.rootId == eles.rootId))
            {
                EleInfo temp;
                map.TryRemove(item.Key, out temp);
            }
            string v = "";

        }
        public static string GetRuntimeIdStr(Array runtimeId)
        {
            string id = string.Empty;
            int index = 0;
            while (index <= runtimeId.Length-1)
            {
                id += runtimeId.GetValue(index).ToString();
                index++;
                if (index <= runtimeId.Length - 1) id += ",";
            }
            return id;
        }
        public static List<EleInfo> GetAllElementEx(IUIAutomationElement root, IUIAutomationElement ele, int level)
        {
            List<EleInfo> eles = new List<EleInfo>();
            CUIAutomation uia = new CUIAutomation();
            IUIAutomationPropertyCondition find1_condition =
                   uia.CreatePropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId,
                    UIA_ControlTypeIds.UIA_WindowControlTypeId) as IUIAutomationPropertyCondition;
            

            IUIAutomationElementArray arry = ele.FindAll(TreeScope.TreeScope_Children, (new CUIAutomation()).CreateTrueCondition());
            for (int i = 0; i < arry.Length; i++)
            {
                eles.Add(new EleInfo(root, arry.GetElement(i), level));
            }
            
            return eles;
        }

        private static Dictionary<string, object> GetElementInfo(IUIAutomationElement element)
        {
            var dictResult = new Dictionary<string, object>();
            if (element == null)
            {
                return dictResult;
            }

            dictResult["AcceleratorKey"] = element.CurrentAcceleratorKey;
            dictResult["AccessKey"] = element.CurrentAccessKey;
            dictResult["AriaProperties"] = element.CurrentAriaProperties;
            dictResult["AriaRole"] = element.CurrentAriaRole;
            dictResult["AutomationId"] = element.CurrentAutomationId;
            dictResult["BoundingRectangle"] = element.CurrentBoundingRectangle;
            dictResult["ClassName"] = element.CurrentClassName;
            dictResult["ControlType"] = element.CurrentControlType;
            dictResult["Culture"] = element.CurrentCulture;
            dictResult["FrameworkId"] = element.CurrentFrameworkId;
            dictResult["HasKeyboardFocus"] = element.CurrentHasKeyboardFocus;
            dictResult["HelpText"] = element.CurrentHelpText;
            dictResult["IsEnabled"] = element.CurrentIsEnabled;
            dictResult["IsKeyboardFocusable"] = element.CurrentIsKeyboardFocusable;
            dictResult["IsOffscreen"] = element.CurrentIsOffscreen > 0;
            dictResult["IsPassword"] = element.CurrentIsPassword > 0;
            dictResult["IsContentElement"] = element.CurrentIsContentElement > 0;
            dictResult["IsControlElement"] = element.CurrentIsControlElement > 0;
            dictResult["ItemStatus"] = element.CurrentItemStatus;
            dictResult["ItemType"] = element.CurrentItemType;
            dictResult["LocalizedControlType"] = element.CurrentLocalizedControlType;
            dictResult["Name"] = element.CurrentName;
            dictResult["ProcessId"] = element.CurrentProcessId;
            dictResult["ProviderDescription"] = element.CurrentProviderDescription;

            return dictResult;
        }
        public static bool SelectComboItem(IUIAutomationElement control)
        {
            if (control == null)
            {
                return false;
            }

            object pattern = control.GetCurrentPattern(UIA_PatternIds.UIA_SelectionItemPatternId);
            if (pattern == null)
            {
                return false;
            }

            IUIAutomationSelectionItemPattern select_item_pattern = pattern as IUIAutomationSelectionItemPattern;
            select_item_pattern.Select();

            object invoke_pattern = control.GetCurrentPattern(UIA_PatternIds.UIA_InvokePatternId);
            if (invoke_pattern != null)
            {
                IUIAutomationInvokePattern invoke = invoke_pattern as IUIAutomationInvokePattern;
                invoke.Invoke();
            }

            return true;
        }

        public static bool CompareRuntimeId(
            Array array1,
            Array array2)
        {
            if (array1.Length != array2.Length)
            {
                return false;
            }

            int last_index = array1.Length - 1;
            while (last_index > 0)
            {
                if (!array1.GetValue(last_index).ToString().Equals(array2.GetValue(last_index).ToString()))
                {
                    return false;
                }
                last_index--;
            }

            return true;
        }


        #endregion

        #region Private


        private static IUIAutomationElement FindControl(
            IUIAutomationTreeWalker treeWalker,
            IUIAutomationElement element,
            string automationid,
            string ctrlname,
            Array runtimeid)
        {
            IUIAutomationElement control = null;
            if (element == null || treeWalker == null)
            {
                return control;
            }

            while (element != null && control == null)
            {
                if (!string.IsNullOrEmpty(automationid) && !string.IsNullOrEmpty(ctrlname) &&
                    !string.IsNullOrEmpty(element.CurrentName) && !string.IsNullOrEmpty(element.CurrentAutomationId) &&
                    element.CurrentName.Equals(ctrlname, StringComparison.OrdinalIgnoreCase) &&
                    element.CurrentAutomationId.Equals(automationid, StringComparison.OrdinalIgnoreCase))
                {
                    control = element;
                    break;
                }
                else if (string.IsNullOrEmpty(ctrlname) && !string.IsNullOrEmpty(element.CurrentAutomationId) &&
                    element.CurrentAutomationId.Equals(automationid, StringComparison.OrdinalIgnoreCase))
                {
                    control = element;
                    break;
                }
                else if (string.IsNullOrEmpty(automationid) && !string.IsNullOrEmpty(element.CurrentName) &&
                    element.CurrentName.Equals(ctrlname, StringComparison.OrdinalIgnoreCase))
                {
                    control = element;
                    break;
                }
                else if (runtimeid != null)
                {
                    Array curRuntimeId = element.GetRuntimeId();
                    if (CompareRuntimeId(curRuntimeId, runtimeid))
                    {
                        control = element;
                        break;
                    }
                }

                IUIAutomationElement subElement = treeWalker.GetFirstChildElement(element);
                if (subElement != null)
                {
                    control = FindControl(treeWalker, subElement, automationid, ctrlname, runtimeid);
                }

                if (control == null)
                {
                    element = treeWalker.GetNextSiblingElement(element);
                }
            }

            return control;
        }
        #endregion
    }
}
