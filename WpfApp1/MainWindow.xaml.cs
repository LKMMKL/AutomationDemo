using System;
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
            //Task.Run(() =>
            //{
            //    Thread.Sleep(3000);
            //    var target = GetTreeViewItemFromObject(this.tree.ItemContainerGenerator, list[2]);
            //    this.Dispatcher.Invoke(() =>
            //    {
            //        target.IsSelected = true;
            //        target.IsExpanded= true;
            //    });
                
            //});
        }

        private void tree_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            TreeViewItem item = sender as TreeViewItem;

        }

        private void tree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            TreeView treeView = sender as TreeView;
            EleInfo item = (EleInfo)treeView.SelectedItem;
     
            this.eleName.Content = item.Name;
            this.eleClassName.Content = item.ClassName;
            this.eleId.Content = item.AutomationId;
            tagRECT rect = item.curr.CurrentBoundingRectangle;
            HightLight.DrawHightLight(rect);
        }

        private TreeViewItem GetTreeViewItemFromObject(ItemContainerGenerator container, object tvio)
        {
            var item = container.ContainerFromItem(tvio) as TreeViewItem;
            if (item != null)
            {
                return item;
            }

            for (int i = 0; i < container.Items.Count; i++)
            {
                var subContainer = (TreeViewItem)container.ContainerFromIndex(i);
                if (subContainer != null)
                {
                    item = GetTreeViewItemFromObject(subContainer.ItemContainerGenerator, tvio);
                    if (item != null)
                    {
                        return item;
                    }
                }
            }

            return null;
        }

        public void MouseSelect(object ele)
        {
            var target = GetTreeViewItemFromObject(this.tree.ItemContainerGenerator, ele);
            if (target != null)
            {
                this.Dispatcher.Invoke(() =>
                {
                    target.IsSelected = true;
                    target.IsExpanded = true;
                });
            }
            
        }
    }

    public class RelayCommand : ICommand
    {
        public event EventHandler CanExecuteChanged;

        private Action mExecute;
        public event EventHandler canExecuteChanged;

        public RelayCommand(Action execute)
        {
            mExecute = execute;
        }
        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            mExecute?.Invoke();
        }
    }
}
