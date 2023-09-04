using System;
using System.CodeDom;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using UIAutomationClient;
using WpfApp1.Util;

namespace WpfApp1
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {

        public event PropertyChangedEventHandler PropertyChanged;

        public static List<EleInfo> list { get; set; }

        public static EleInfo currItem { get; set; }

        public MainWindow()
        {
            InitializeComponent();

            this.DataContext = this;
            list = UIControlAssist.GetAllElement();
            HightLight.mouseFunc = MouseSelect;
           
        }

        private void tree_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            TreeViewItem item = sender as TreeViewItem;

        }

        private void tree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            TreeView treeView = sender as TreeView;
            EleInfo item = (EleInfo)treeView.SelectedItem;
            if (item == null) return;
            this.nodeName.Content = item.name;
            this.nodeClassName.Content = item.className;
            this.nodeAutomationid.Content = item.automationId;
            this.nodeRuntimeid.Content = item.runtimeId;
            this.nodeRect.Content = item.rect;
            this.nodeType.Content = item.type;
            this.nodeOffScreen.Content = item.offScreen;
            tagRECT rect = item.curr.CurrentBoundingRectangle;
            HightLight.DrawHightLight(rect);
        }

        private TreeViewItem GetTreeViewItem(ItemsControl container, object item)
        {
            if (container != null)
            {
                if (container.DataContext == item)
                {
                    return container as TreeViewItem;
                }

                // Expand the current container
                if (container is TreeViewItem && !((TreeViewItem)container).IsExpanded)
                {
                    container.SetValue(TreeViewItem.IsExpandedProperty, true);
                }

                // Try to generate the ItemsPresenter and the ItemsPanel.
                // by calling ApplyTemplate.  Note that in the
                // virtualizing case even if the item is marked
                // expanded we still need to do this step in order to
                // regenerate the visuals because they may have been virtualized away.

                container.ApplyTemplate();
                ItemsPresenter itemsPresenter =
                    (ItemsPresenter)container.Template.FindName("ItemsHost", container);
                if (itemsPresenter != null)
                {
                    itemsPresenter.ApplyTemplate();
                }
                else
                {
                    // The Tree template has not named the ItemsPresenter,
                    // so walk the descendents and find the child.
                    itemsPresenter = FindVisualChild<ItemsPresenter>(container);
                    if (itemsPresenter == null)
                    {
                        container.UpdateLayout();

                        itemsPresenter = FindVisualChild<ItemsPresenter>(container);
                    }
                }

                Panel itemsHostPanel = (Panel)VisualTreeHelper.GetChild(itemsPresenter, 0);

                // Ensure that the generator for this panel has been created.
                UIElementCollection children = itemsHostPanel.Children;

                for (int i = 0, count = container.Items.Count; i < count; i++)
                {
                    TreeViewItem subContainer;
                    subContainer =
                            (TreeViewItem)container.ItemContainerGenerator.
                            ContainerFromIndex(i);

                    // Bring the item into view to maintain the
                    // same behavior as with a virtualizing panel.
                    subContainer.BringIntoView();

                    if (subContainer != null)
                    {
                        // Search the next level for the object.
                        TreeViewItem resultContainer = GetTreeViewItem(subContainer, item);
                        if (resultContainer != null)
                        {
                            return resultContainer;
                        }
                        else
                        {
                            // The object is not under this TreeViewItem
                            // so collapse it.
                            subContainer.IsExpanded = false;
                        }
                    }
                }
            }

            return null;
        }
        private T FindVisualChild<T>(Visual visual) where T : Visual
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(visual); i++)
            {
                Visual child = (Visual)VisualTreeHelper.GetChild(visual, i);
                if (child != null)
                {
                    T correctlyTyped = child as T;
                    if (correctlyTyped != null)
                    {
                        return correctlyTyped;
                    }

                    T descendent = FindVisualChild<T>(child);
                    if (descendent != null)
                    {
                        return descendent;
                    }
                }
            }

            return null;
        }
        public void MouseSelect(object ele)
        {
            //foreach (var item in tree.Items)
            //{
            //    DependencyObject dObject = tree.ItemContainerGenerator.ContainerFromItem(item);
            //    ((TreeViewItem)dObject).ExpandSubtree();
            //}
            var target = GetTreeViewItem(tree, ele);
            if (target != null)
            {
                this.Dispatcher.Invoke(() =>
                {
                    target.IsSelected = true;
                    target.IsExpanded = true;
                });
            }
            
        }

        private void Refresh(object sender, RoutedEventArgs e)
        {
            var runtimeId = ((Button)sender).Tag;
            EleInfo eles = (EleInfo)tree.Items[0];
            var index = 0;
           
            for(int i = 0; i < eles.childs.Count; i++)
            {
                index = i;
                if(eles.childs[i].runtimeId == runtimeId)
                {
                    
                    break;
                }

            }
            EleInfo target = eles.childs[index];
            UIControlAssist.Refresh(target);
            Task.Run(() => {
                target = new EleInfo(target.curr, target.curr, 0);
            }).Wait();
            
            
            if (target.curr != null)
            {
                eles.childs[index] = target;
            }
            else
            {
                eles.childs.RemoveAt(index);
            }
            
            //eles.childs.RemoveAll((ele) => ele.runtimeId == runtimeId);
            tree.Items.Refresh();
            TreeViewItem ti = tree.ItemContainerGenerator.ContainerFromIndex(0) as TreeViewItem;
            ti.IsExpanded = true;
        }
    }
}
