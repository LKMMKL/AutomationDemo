using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using UIAutomationClient;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using Point = System.Drawing.Point;

namespace WpfApp1
{
    public class EleInfo
    {
        public string Name { get; set; }
        public string ClassName { get; set; }
        public string AutomationId { get; set; }
        public tagRECT rect { get; set; }

        public int level = 0;
        public List<EleInfo> childs { get; set; }

        public IUIAutomationElement root; // 根窗口
        public IUIAutomationElement parent; // 父节点
        public IUIAutomationElement curr; //当前节点
        public EleInfo()
        {
            
        }

        public EleInfo(IUIAutomationElement parent, IUIAutomationElement ele, int level=0)
        {
            try
            {
                Name = ele.CurrentName == string.Empty ? "\"\"" : ele.CurrentName;
                ClassName = ele.CurrentClassName;
                rect = ele.CurrentBoundingRectangle;
                AutomationId = ele.CurrentAutomationId;
                this.parent = parent;
                curr = ele;
                this.level = ++level;
                if (this.level < 4)
                {
                    childs = UIControlAssist.GetAllElementEx(curr, this.level);
                }
                else
                {
                    // 空元素
                    childs = new List<EleInfo>() { new EleInfo() };
                }
                //childs = UIControlAssist.GetAllElementEx(curr, this.level);
            }
            catch(COMException ce) {
            }
            

        }

        public void Retrive()
        {
            childs = UIControlAssist.GetAllElementEx(curr, this.level);
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
            CountdownEvent countdownEvent = new CountdownEvent(arry.Length);
            for (int i = 0; i < arry.Length; i++)
            {
                var index = i;
                Task.Run(() =>
                {
                    eles.Add(new EleInfo(uia.GetRootElement(), arry.GetElement(index)));
                    countdownEvent.Signal();
                });
            }
            countdownEvent.Wait();
            allEle = eles;
            return eles;
        }

        public static List<EleInfo> GetAllElementEx(IUIAutomationElement ele, int level)
        {
            List<EleInfo> eles = new List<EleInfo>();
            CUIAutomation uia = new CUIAutomation();
            IUIAutomationPropertyCondition find1_condition =
                   uia.CreatePropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId,
                    UIA_ControlTypeIds.UIA_WindowControlTypeId) as IUIAutomationPropertyCondition;
            

            IUIAutomationElementArray arry = ele.FindAll(TreeScope.TreeScope_Children, (new CUIAutomation()).CreateTrueCondition());
            for (int i = 0; i < arry.Length; i++)
            {
                eles.Add(new EleInfo(ele, arry.GetElement(i), level));
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

        public static bool IsSubControlExist(
            IUIAutomationElement control,
            string automationid,
            string ctrlname)
        {
            if (null == control)
            {
                return false;
            }

            CUIAutomation uia = new CUIAutomation();
            IUIAutomationElement subcontrol = null;
            if (!string.IsNullOrEmpty(ctrlname))
            {
                subcontrol = control.FindFirst(TreeScope.TreeScope_Descendants
                    , uia.CreatePropertyConditionEx(UIA_PropertyIds.UIA_NamePropertyId,
                    ctrlname, PropertyConditionFlags.PropertyConditionFlags_IgnoreCase));
            }
            else if (!string.IsNullOrEmpty(automationid))
            {
                subcontrol = control.FindFirst(TreeScope.TreeScope_Descendants
                    , uia.CreatePropertyConditionEx(UIA_PropertyIds.UIA_AutomationIdPropertyId,
                    automationid, PropertyConditionFlags.PropertyConditionFlags_IgnoreCase));
            }
            if (subcontrol == null)
            {
                return false;
            }

            IUIAutomationLegacyIAccessiblePattern iAccessible = subcontrol.GetCurrentPattern(UIA_PatternIds.UIA_LegacyIAccessiblePatternId) as IUIAutomationLegacyIAccessiblePattern;
            if (iAccessible == null)
            {
                return false;
            }

            uint State = iAccessible.CurrentState;

            return 0x8000 != (State & 0x8000);
        }

        public static void InitScrollPosition(
            IUIAutomationScrollPattern scrollpattern,
            bool isvertical)
        {
            double scroll_percent = 0.0;
            do
            {
                scroll_percent = isvertical ? scrollpattern.CurrentVerticalScrollPercent : scrollpattern.CurrentHorizontalScrollPercent;
                if (scrollpattern.CurrentVerticallyScrollable > 0 && isvertical && scroll_percent > 0.0)
                {
                    scrollpattern.Scroll(ScrollAmount.ScrollAmount_NoAmount, ScrollAmount.ScrollAmount_LargeDecrement);
                }
                else if (scrollpattern.CurrentHorizontallyScrollable > 0 && scroll_percent > 0.0)
                {
                    scrollpattern.Scroll(ScrollAmount.ScrollAmount_LargeDecrement, ScrollAmount.ScrollAmount_NoAmount);
                }
            } while (scroll_percent > 0.0);
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

        public static IUIAutomationElement GetListItem(
            IUIAutomationElement control,
            int row)
        {
            IUIAutomationElement element = null;
            if (row < 0)
            {
                return element;
            }

            IUIAutomation uia = new CUIAutomation();
            IUIAutomationPropertyCondition condition = uia.CreatePropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_ListItemControlTypeId) as IUIAutomationPropertyCondition;
            IUIAutomationElementArray item_element_array = control.FindAll(TreeScope.TreeScope_Subtree, condition);
            if (item_element_array.Length > row)
            {
                element = item_element_array.GetElement(row);
            }

            return element;
        }

        public static IUIAutomationElement GetMainWindow(
            string classname,
            string windowcaption)
        {
            IUIAutomationElement window = GetMainWindowEx(classname, windowcaption);
            if (window != null)
            {
                string msg = "The window is success to find 。The window's info [ windowcaption: " +
                             window.CurrentName + ",classname: " + window.CurrentClassName +
                             ",automationid: " + window.CurrentAutomationId + "]";
                return window;
            }
            return window;
        }

        public static IUIAutomationElement GetMainWindowEx(
            string classname,
            string windowcaption)
        {
            IUIAutomation uia = new CUIAutomation();
            IUIAutomationElement root = uia.GetRootElement();

            IUIAutomationCondition find_condition = null;
            if (!string.IsNullOrEmpty(classname) && !string.IsNullOrEmpty(windowcaption))
            {
                IUIAutomationPropertyCondition find1_condition =
                    uia.CreatePropertyConditionEx(UIA_PropertyIds.UIA_NamePropertyId, windowcaption,
                    PropertyConditionFlags.PropertyConditionFlags_IgnoreCase) as IUIAutomationPropertyCondition;

                IUIAutomationPropertyCondition find2_condition =
                    uia.CreatePropertyConditionEx(UIA_PropertyIds.UIA_ClassNamePropertyId, classname,
                    PropertyConditionFlags.PropertyConditionFlags_IgnoreCase) as IUIAutomationPropertyCondition;

                find_condition = uia.CreateAndCondition(find1_condition, find2_condition);
            }
            else if (!string.IsNullOrEmpty(windowcaption))
            {
                find_condition = uia.CreatePropertyConditionEx(UIA_PropertyIds.UIA_NamePropertyId, windowcaption,
                    PropertyConditionFlags.PropertyConditionFlags_IgnoreCase) as IUIAutomationPropertyCondition;
            }
            else if (!string.IsNullOrEmpty(classname))
            {
                find_condition = uia.CreatePropertyConditionEx(UIA_PropertyIds.UIA_ClassNamePropertyId, classname,
                    PropertyConditionFlags.PropertyConditionFlags_IgnoreCase) as IUIAutomationPropertyCondition;
            }

            IUIAutomationPropertyCondition type1_condition =
                uia.CreatePropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId,
                UIA_ControlTypeIds.UIA_WindowControlTypeId) as IUIAutomationPropertyCondition;

            IUIAutomationPropertyCondition type2_condition =
                uia.CreatePropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId,
                UIA_ControlTypeIds.UIA_PaneControlTypeId) as IUIAutomationPropertyCondition;

            IUIAutomationPropertyCondition type3_condition =
                uia.CreatePropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId,
                UIA_ControlTypeIds.UIA_MenuControlTypeId) as IUIAutomationPropertyCondition;

            IUIAutomationPropertyCondition type4_condition =
                uia.CreatePropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId,
                UIA_ControlTypeIds.UIA_ListControlTypeId) as IUIAutomationPropertyCondition;

            IUIAutomationCondition type_or_1_condition = uia.CreateOrCondition(type1_condition, type2_condition);
            IUIAutomationCondition type_or_2_condition = uia.CreateOrCondition(type3_condition, type4_condition);
            IUIAutomationCondition type_condition = uia.CreateOrCondition(type_or_1_condition, type_or_2_condition);

            IUIAutomationCondition and_condition = uia.CreateAndCondition(find_condition, type_condition);
            return root.FindFirst(TreeScope.TreeScope_Subtree, and_condition);
        }

        public static IUIAutomationElement GetControl(
            string classname,
            string windowcaption,
            string automationid,
            string ctrlname,
            Array runtimeid)
        {
            //return GetControl(GetMainWindow(classname, windowcaption), automationid, ctrlname);
            IUIAutomationElement control = null;
            try
            {
                control = GetControlEx(GetMainWindow(classname, windowcaption), automationid, ctrlname, runtimeid);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                control = GetControlEx(GetMainWindow(classname, windowcaption), automationid, ctrlname, runtimeid);
            }

            return control;
        }

        public static IUIAutomationElement GetControl(
            IUIAutomationElement parent,
            string automationid,
            string ctrlname)
        {
            IUIAutomationElement control = null;
            if (parent == null)
            {
                return control;
            }

            IUIAutomation uia = new CUIAutomation();
            IUIAutomationPropertyCondition condition = null;
            if (!string.IsNullOrEmpty(automationid))
            {
                condition = uia.CreatePropertyConditionEx(UIA_PropertyIds.UIA_AutomationIdPropertyId,
                    automationid, PropertyConditionFlags.PropertyConditionFlags_IgnoreCase
                    ) as IUIAutomationPropertyCondition;
            }
            else if (!string.IsNullOrEmpty(ctrlname))
            {
                condition = uia.CreatePropertyConditionEx(UIA_PropertyIds.UIA_NamePropertyId,
                    ctrlname, PropertyConditionFlags.PropertyConditionFlags_IgnoreCase
                    ) as IUIAutomationPropertyCondition;
            }

            if (condition != null)
            {
                control = parent.FindFirst(TreeScope.TreeScope_Subtree, condition);
            }

            return control;
        }

        public static IUIAutomationElement GetControlEx(
            IUIAutomationElement parent,
            string automationid,
            string ctrlname,
            Array runtimeid)
        {
            IUIAutomationElement control = null;
            if (parent == null)
            {
                return control;
            }

            IUIAutomation uia = new CUIAutomation();
            IUIAutomationCondition condition = uia.CreateTrueCondition();
            IUIAutomationTreeWalker treeWalker = uia.CreateTreeWalker(condition);
            IUIAutomationElement element = treeWalker.GetFirstChildElement(parent);
            if (element != null && treeWalker != null)
            {
                control = FindControl(treeWalker, element, automationid, ctrlname, runtimeid);
            }

            if (control != null)
            {
                string msg = "The control is success to find 。The control's info [ ctrlname: " +
                             control.CurrentName + ",classname: " + control.CurrentClassName +
                             ",automationid: " + control.CurrentAutomationId + "]";
            }
            else
            {
                string msg = "The control is fail to find ！What you find is [ ctrlname: " +
                             ctrlname + ",automationid: " + automationid + "]";
            }
            return control;
        }

        public static List<object> GetListItemText(
            IUIAutomationElement control,
            int row)
        {
            List<object> result_list = new List<object>();
            if (control == null)
            {
                return result_list;
            }

            IUIAutomation uia = new CUIAutomation();
            IUIAutomationPropertyCondition condition = uia.CreatePropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_ListItemControlTypeId) as IUIAutomationPropertyCondition;
            IUIAutomationElementArray item_element_array = control.FindAll(UIAutomationClient.TreeScope.TreeScope_Children, condition);
            for (int j = 0; j < item_element_array.Length; j++)
            {
                int index = j;
                if (row != -1)
                {
                    index = row;
                }

                List<string> item_contents_list = new List<string>();
                IUIAutomationPropertyCondition header_item_condition = uia.CreatePropertyCondition(UIA_PropertyIds.UIA_ControlTypePropertyId, UIA_ControlTypeIds.UIA_TextControlTypeId) as IUIAutomationPropertyCondition;
                IUIAutomationTreeWalker tree_walker = uia.CreateTreeWalker(header_item_condition);
                IUIAutomationElement parent_element = item_element_array.GetElement(index);
                IUIAutomationElement element = tree_walker.GetFirstChildElement(parent_element);
                IUIAutomationElement last_element = tree_walker.GetLastChildElement(item_element_array.GetElement(index));
                if (last_element != null)
                {
                    Array last_element_array = last_element.GetRuntimeId();
                    while (element != null &&
                        last_element != null)
                    {
                        item_contents_list.Add(element.CurrentName);
                        Array element_array = element.GetRuntimeId();
                        if (CompareRuntimeId(last_element_array, element_array))
                        {
                            break;
                        }

                        element = tree_walker.GetNextSiblingElement(element);
                    }
                }
                else
                {
                    item_contents_list.Add(parent_element.CurrentName);
                }

                result_list.Add(item_contents_list);
                if (row != -1)
                {
                    break;
                }
            }

            return result_list;
        }

        public static string GetListItemContents(IUIAutomationElement control, int row)
        {
            //string result = string.Empty;
            //if (control != null)
            //{
            //    List<object> items_contents_list = GetListItemText(control, row);
            //    return JsonConvert.SerializeObject(items_contents_list);
            //}
            //return result;
            return "";
        }

        public static List<string> GetComboBoxItemsList(IUIAutomationElement element)
        {
            List<string> _result = new List<string>();
            try
            {
                object invokePattern = null;
                if (element.CurrentIsEnabled > 0 &&
                    TryGetCurrentPattern(element, UIA_PatternIds.UIA_ExpandCollapsePatternId, out invokePattern))
                {
                    IUIAutomationExpandCollapsePattern expandCollapsePattern = invokePattern as IUIAutomationExpandCollapsePattern;
                    expandCollapsePattern.Expand();
                    IUIAutomationElementArray subElementArray = element.FindAll(TreeScope.TreeScope_Children, (new CUIAutomation()).CreateTrueCondition());
                    for (int i = 0; i < subElementArray.Length; i++)
                    {
                        IUIAutomationElement subElement = subElementArray.GetElement(i);
                        if (subElement.CurrentControlType == UIA_ControlTypeIds.UIA_ListItemControlTypeId)
                        {
                            _result.Add(GetListItemName(subElement));
                        }
                        else if (subElement.CurrentControlType == UIA_ControlTypeIds.UIA_ListControlTypeId)
                        {
                            string items = GetListItemContents(subElement, -1);
                            //JToken jt = JToken.Parse(items);
                            //if (jt != null)
                            //{
                            //    for (int j = 0; j < jt.Count(); j++)
                            //    {
                            //        _result.Add(jt[j][0].ToString());
                            //    }
                            //}
                        }
                    }
                    expandCollapsePattern.Collapse();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return _result;
        }

        public static List<string> GetComboItemName(IUIAutomationElement control)
        {
            List<string> result = new List<string>(); ;

            IUIAutomationElementArray elementCollection = GetComboItemCollection(control);
            if (elementCollection == null) return result;
            if (elementCollection.Length == 0) return result;

            result = new List<string>();
            for (int i = 0; i < elementCollection.Length; i++)
            {
                IUIAutomationElement element = elementCollection.GetElement(i);
                if (element.CurrentControlType != UIA_ControlTypeIds.UIA_ListItemControlTypeId) continue;
                result.Add(GetComboItemNameSub(element));
            }

            return result;
        }

        public static IUIAutomationElement GetRawViewWalkerParentElement(IUIAutomationElement element)
        {
            try
            {
                return (new CUIAutomation()).RawViewWalker.GetParentElement(element);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static IUIAutomationElementArray GetChildrenAllElements(IUIAutomationElement element)
        {
            return element.FindAll(TreeScope.TreeScope_Children, (new CUIAutomation()).CreateTrueCondition());
        }

        public static bool TryGetCurrentPattern(IUIAutomationElement element,
            int patternId,
            out object pattern)
        {
            pattern = null;
            try
            {
                pattern = element.GetCurrentPattern(patternId);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return true;
        }

        public static string GetDataItemTexts(IUIAutomationElement control)
        {
            StringBuilder result = new StringBuilder();

            IUIAutomationElementArray children = GetChildrenAllElements(control);

            for (int i = 0; i < children.Length; i++)
            {
                IUIAutomationElement child = children.GetElement(i);
                if (child.CurrentControlType == UIA_ControlTypeIds.UIA_TextControlTypeId)
                {
                    result.Append(" " + GetControlName(child));
                }
            }
            return result.ToString();
        }

        public static string GetControlName(IUIAutomationElement control)
        {
            string result = string.Empty;
            if (control != null)
            {
                result = control.CurrentName;
            }

            return result;
        }

        

        public static Dictionary<string, int> GetPosition(
            int left,
            int right,
            int top,
            int bottom,
            ControlScope scope)
        {
            Dictionary<string, int> dictPoint = new Dictionary<string, int>();
            switch (scope)
            {
                case ControlScope.Bottom:
                    {
                        dictPoint["x"] = (left + right) / 2;
                        dictPoint["y"] = bottom - ((bottom - top) / 20);
                        break;
                    }
                case ControlScope.Central:
                    {
                        dictPoint["x"] = (left + right) / 2;
                        dictPoint["y"] = (bottom + top) / 2;
                        break;
                    }
                case ControlScope.Left:
                    {
                        dictPoint["x"] = left + (right - left) / 20;
                        dictPoint["y"] = (bottom + top) / 2;
                        break;
                    }
                case ControlScope.Right:
                    {
                        dictPoint["x"] = right - (right - left) / 20;
                        dictPoint["y"] = (bottom + top) / 2;
                        break;
                    }
                case ControlScope.Top:
                    {
                        dictPoint["x"] = (left + right) / 2;
                        dictPoint["y"] = top + ((bottom - top) / 20);
                        break;
                    }
                default:
                    break;
            }

            return dictPoint;
        }

        public static List<string> GetControlState(IUIAutomationElement control)
        {
            List<string> listState = new List<string>();
            if (control != null)
            {
                object pattern = control.GetCurrentPattern(UIA_PatternIds.UIA_LegacyIAccessiblePatternId);
                if (pattern != null)
                {
                    IUIAutomationLegacyIAccessiblePattern legacy_i_accessible_pattern =
                        pattern as IUIAutomationLegacyIAccessiblePattern;
                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_ALERT_HIGH)
                        == ControlState.STATE_SYSTEM_ALERT_HIGH)
                        listState.Add(ControlState.STATE_SYSTEM_ALERT_HIGH.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_ALERT_LOW)
                        == ControlState.STATE_SYSTEM_ALERT_LOW)
                        listState.Add(ControlState.STATE_SYSTEM_ALERT_LOW.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_ALERT_MEDIUM)
                        == ControlState.STATE_SYSTEM_ALERT_MEDIUM)
                        listState.Add(ControlState.STATE_SYSTEM_ALERT_MEDIUM.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_ANIMATED)
                        == ControlState.STATE_SYSTEM_ANIMATED)
                        listState.Add(ControlState.STATE_SYSTEM_ANIMATED.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_BUSY)
                        == ControlState.STATE_SYSTEM_BUSY)
                        listState.Add(ControlState.STATE_SYSTEM_BUSY.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_CHECKED)
                        == ControlState.STATE_SYSTEM_CHECKED)
                        listState.Add(ControlState.STATE_SYSTEM_CHECKED.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_COLLAPSED)
                        == ControlState.STATE_SYSTEM_COLLAPSED)
                        listState.Add(ControlState.STATE_SYSTEM_COLLAPSED.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_DEFAULT)
                        == ControlState.STATE_SYSTEM_DEFAULT)
                        listState.Add(ControlState.STATE_SYSTEM_DEFAULT.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_EXPANDED)
                        == ControlState.STATE_SYSTEM_EXPANDED)
                        listState.Add(ControlState.STATE_SYSTEM_EXPANDED.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_EXTSELECTABLE)
                        == ControlState.STATE_SYSTEM_EXTSELECTABLE)
                        listState.Add(ControlState.STATE_SYSTEM_EXTSELECTABLE.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_FLOATING)
                        == ControlState.STATE_SYSTEM_FLOATING)
                        listState.Add(ControlState.STATE_SYSTEM_FLOATING.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_FOCUSABLE)
                        == ControlState.STATE_SYSTEM_FOCUSABLE)
                        listState.Add(ControlState.STATE_SYSTEM_FOCUSABLE.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_FOCUSED)
                        == ControlState.STATE_SYSTEM_FOCUSED)
                        listState.Add(ControlState.STATE_SYSTEM_FOCUSED.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_HOTTRACKED)
                        == ControlState.STATE_SYSTEM_HOTTRACKED)
                        listState.Add(ControlState.STATE_SYSTEM_HOTTRACKED.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_INDETERMINATE)
                        == ControlState.STATE_SYSTEM_INDETERMINATE)
                        listState.Add(ControlState.STATE_SYSTEM_INDETERMINATE.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_INVISIBLE)
                        == ControlState.STATE_SYSTEM_INVISIBLE)
                        listState.Add(ControlState.STATE_SYSTEM_INVISIBLE.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_LINKED)
                        == ControlState.STATE_SYSTEM_LINKED)
                        listState.Add(ControlState.STATE_SYSTEM_LINKED.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_MARQUEED)
                        == ControlState.STATE_SYSTEM_MARQUEED)
                        listState.Add(ControlState.STATE_SYSTEM_MARQUEED.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_MIXED)
                        == ControlState.STATE_SYSTEM_MIXED)
                        listState.Add(ControlState.STATE_SYSTEM_MIXED.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_MOVEABLE)
                        == ControlState.STATE_SYSTEM_MOVEABLE)
                        listState.Add(ControlState.STATE_SYSTEM_MOVEABLE.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_MULTISELECTABLE)
                        == ControlState.STATE_SYSTEM_MULTISELECTABLE)
                        listState.Add(ControlState.STATE_SYSTEM_MULTISELECTABLE.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_OFFSCREEN)
                        == ControlState.STATE_SYSTEM_OFFSCREEN)
                        listState.Add(ControlState.STATE_SYSTEM_OFFSCREEN.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_PRESSED)
                        == ControlState.STATE_SYSTEM_PRESSED)
                        listState.Add(ControlState.STATE_SYSTEM_PRESSED.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_PROTECTED)
                        == ControlState.STATE_SYSTEM_PROTECTED)
                        listState.Add(ControlState.STATE_SYSTEM_PROTECTED.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_READONLY)
                        == ControlState.STATE_SYSTEM_READONLY)
                        listState.Add(ControlState.STATE_SYSTEM_READONLY.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_SELECTABLE)
                        == ControlState.STATE_SYSTEM_SELECTABLE)
                        listState.Add(ControlState.STATE_SYSTEM_SELECTABLE.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_SELECTED)
                        == ControlState.STATE_SYSTEM_SELECTED)
                        listState.Add(ControlState.STATE_SYSTEM_SELECTED.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_SELFVOICING)
                        == ControlState.STATE_SYSTEM_SELFVOICING)
                        listState.Add(ControlState.STATE_SYSTEM_SELFVOICING.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_SIZEABLE)
                        == ControlState.STATE_SYSTEM_SIZEABLE)
                        listState.Add(ControlState.STATE_SYSTEM_SIZEABLE.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_TRAVERSED)
                        == ControlState.STATE_SYSTEM_TRAVERSED)
                        listState.Add(ControlState.STATE_SYSTEM_TRAVERSED.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_UNAVAILABLE)
                        == ControlState.STATE_SYSTEM_UNAVAILABLE)
                        listState.Add(ControlState.STATE_SYSTEM_UNAVAILABLE.ToString());

                    if (((ControlState)legacy_i_accessible_pattern.CurrentState & ControlState.STATE_SYSTEM_VALID)
                        == ControlState.STATE_SYSTEM_VALID)
                        listState.Add(ControlState.STATE_SYSTEM_VALID.ToString());
                }
            }

            return listState;
        }

        public static bool Scroll(
            IUIAutomationElement control,
            IUIAutomationScrollPattern scrollpattern,
            string find_automatonid,
            string find_ctrlname)
        {
            InitScrollPosition(scrollpattern, true);
            bool result = false;
            double v_scroll_percent = 0.0;
            double h_scroll_percent = 0.0;
            do
            {
                InitScrollPosition(scrollpattern, false);
                while (scrollpattern.CurrentHorizontallyScrollable > 0 && h_scroll_percent < 99.9 && h_scroll_percent >= 0.0)
                {
                    if (IsSubControlExist(control, find_automatonid, find_ctrlname))
                    {
                        result = true;
                        break;
                    }

                    if (scrollpattern.CurrentHorizontallyScrollable > 0)
                    {
                        h_scroll_percent = scrollpattern.CurrentHorizontalScrollPercent;
                        if (h_scroll_percent < 99.9 && h_scroll_percent >= 0.0)
                        {
                            scrollpattern.Scroll(ScrollAmount.ScrollAmount_LargeIncrement, ScrollAmount.ScrollAmount_NoAmount);
                        }
                    }
                }

                if (result)
                {
                    break;
                }

                if (IsSubControlExist(control, find_automatonid, find_ctrlname))
                {
                    result = true;
                    break;
                }

                if (scrollpattern.CurrentVerticallyScrollable > 0)
                {
                    v_scroll_percent = scrollpattern.CurrentVerticalScrollPercent;
                    if (v_scroll_percent < 99.9 && v_scroll_percent >= 0.0)
                    {
                        scrollpattern.Scroll(ScrollAmount.ScrollAmount_NoAmount, ScrollAmount.ScrollAmount_LargeIncrement);
                    }
                }
                System.Threading.Thread.Sleep(200);
                if (IsSubControlExist(control, find_automatonid, find_ctrlname))
                {
                    result = true;
                    break;
                }
            } while (scrollpattern.CurrentVerticallyScrollable > 0 && v_scroll_percent < 99.9 && v_scroll_percent >= 0.0);

            return result;
        }

        public static bool CheckComboItemName(IUIAutomationElement control, string itemname)
        {
            List<string> itemsList = GetComboBoxItemsList(control);
            if (itemsList == null)
            {
                return false;
            }
            return itemsList.Contains(itemname);
        }

        public static List<string> GetComboItemArray(IUIAutomationElement control)
        {
            return GetComboBoxItemsList(control);
        }

        public static int GetControlNumByClassName(string classname,
          string windowcaption,
          string ctrlClassName)
        {
            IUIAutomation uia = new CUIAutomation();
            IUIAutomationElement window = GetMainWindow(classname, windowcaption);
            if (window == null) return 0;
            return window.FindAll(TreeScope.TreeScope_Descendants, uia.CreatePropertyConditionEx(UIA_PropertyIds.UIA_ClassNamePropertyId, ctrlClassName, PropertyConditionFlags.PropertyConditionFlags_IgnoreCase)).Length;
        }

        public static List<Dictionary<string, object>> GetSubElementInfo(IUIAutomationElement control)
        {
            var listResult = new List<Dictionary<string, object>>();

            IUIAutomation uiAutomation = new CUIAutomation();
            IUIAutomationTreeWalker treeWalker = uiAutomation.CreateTreeWalker(uiAutomation.CreateTrueCondition());
            IUIAutomationElement subControl = treeWalker.GetFirstChildElement(control);
            if (subControl == null)
            {
                return listResult;
            }

            listResult.Add(GetElementInfo(subControl));

            IUIAutomationElement nextSubControl = treeWalker.GetNextSiblingElement(subControl);
            if (nextSubControl == null) return listResult;
            while (nextSubControl != null)
            {
                listResult.Add(GetElementInfo(nextSubControl));
                nextSubControl = treeWalker.GetNextSiblingElement(nextSubControl);
            }

            return listResult;
        }

        public static List<Dictionary<string, object>> GetParentElementInfo(IUIAutomationElement control)
        {
            var listResult = new List<Dictionary<string, object>>();

            IUIAutomation uiAutomation = new CUIAutomation();
            IUIAutomationTreeWalker treeWalker = uiAutomation.CreateTreeWalker(uiAutomation.CreateTrueCondition());
            IUIAutomationElement parentControl = treeWalker.GetParentElement(control);
            if (parentControl != null)
            {
                listResult.Add(GetElementInfo(parentControl));
            }

            return listResult;
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

        private static string GetComboItemNameSub(IUIAutomationElement control)
        {
            if (control == null) return string.Empty;
            string result = control.CurrentName;
            IUIAutomationElementArray elementCollection = GetChildrenAllElements(control);
            if (elementCollection == null) return result;
            if (elementCollection.Length == 0) return result;

            for (int i = 0; i < elementCollection.Length; i++)
            {
                IUIAutomationElement element = elementCollection.GetElement(i);
                string name = element.CurrentName;
                if (name.Equals(string.Empty)) continue;
                return name;
            }
            return result;
        }

        private static IUIAutomationElementArray GetComboItemCollection(IUIAutomationElement element)
        {
            IUIAutomationElementArray result = null;
            if (element == null)
            {
                return result;
            }
            IUIAutomationElement comboBox = null;
            if (element.CurrentControlType == UIA_ControlTypeIds.UIA_ComboBoxControlTypeId)
            {
                comboBox = element;
            }
            else if (element.CurrentControlType == UIA_ControlTypeIds.UIA_ListItemControlTypeId)
            {
                comboBox = GetRawViewWalkerParentElement(element);
            }
            if (comboBox == null) return null;

            object invokePattern = null;
            bool _isControlEnable = comboBox.CurrentIsEnabled == 0 ? false : true;
            if (_isControlEnable
                && TryGetCurrentPattern(comboBox, UIA_PatternIds.UIA_ExpandCollapsePatternId, out invokePattern))
            {
                IUIAutomationExpandCollapsePattern expandCollapsePattern = invokePattern as IUIAutomationExpandCollapsePattern;
                if (expandCollapsePattern == null) return result;

                expandCollapsePattern.Expand();
                result = GetChildrenAllElements(comboBox);
                expandCollapsePattern.Collapse();
            }

            return result;
        }

        /// <summary>
        /// 选中与itemName同名的ListItem组件
        /// </summary>
        /// <param name="control"></param>
        /// <param name="itemName"></param>
        /// <returns></returns>
        private static bool SelectListItemEqualsText(IUIAutomationElement control, string itemName)
        {
            //选中ListItem
            if (control == null || control.CurrentControlType != UIA_ControlTypeIds.UIA_ListItemControlTypeId)
            {
                return false;
            }

            //当ListItem的名字与itemName相同时，直接选中
            bool chkName = control.CurrentName.Equals(itemName, StringComparison.OrdinalIgnoreCase);

            if (chkName)
            {
                if (SelectComboItem(control))
                {
                    return true;
                }
            }

            //当ListItem的子控件名字与itemName相同时，直接选中
            var child = RawViewWalkerGetFirstChildElement(control);
            while (child != null)
            {
                bool chkChildName = child.CurrentName.Equals(itemName, StringComparison.OrdinalIgnoreCase);
                if (chkChildName)
                {
                    if (SelectComboItem(control))
                    {
                        return true;
                    }
                }

                child = RawViewWalkerGetNextSiblingElement(child);
            }

            return false;
        }

        private static IUIAutomationElement RawViewWalkerGetFirstChildElement(IUIAutomationElement control)
        {
            var rawViewWalker = (new CUIAutomation()).RawViewWalker;
            if (rawViewWalker == null)
            {
                return null;
            }

            IUIAutomationElement result = null;
            try
            {
                result = rawViewWalker.GetFirstChildElement(control);
            }
            catch
            {
                result = null;
            }

            return result;
        }

        private static IUIAutomationElement RawViewWalkerGetNextSiblingElement(IUIAutomationElement control)
        {
            var rawViewWalker = (new CUIAutomation()).RawViewWalker;
            if (rawViewWalker == null)
            {
                return null;
            }

            IUIAutomationElement result = null;
            try
            {
                result = rawViewWalker.GetNextSiblingElement(control);
            }
            catch
            {
                result = null;
            }

            return result;
        }

        private static string GetListItemName(IUIAutomationElement control)
        {
            string result = string.Empty;
            if (control != null)
            {
                IUIAutomationElementArray elementArray = control.FindAll(TreeScope.TreeScope_Children, (new CUIAutomation()).CreateTrueCondition());
                if (elementArray.Length > 0)
                {
                    for (int i = 0; i < elementArray.Length; i++)
                    {
                        IUIAutomationElement subElement = elementArray.GetElement(i);
                        result = subElement.CurrentControlType == UIA_ControlTypeIds.UIA_TextControlTypeId ?
                            subElement.CurrentName : control.CurrentName;
                    }
                }
                else
                {
                    result = control.CurrentName;
                }
            }

            return result;
        }
        #endregion
    }
}
