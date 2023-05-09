using Autodesk.Revit.DB;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
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

namespace Branch
{
    /// <summary>
    /// LinkWindow.xaml 的交互逻辑
    /// </summary>
    public sealed partial class LinkWindow : Window
    {

        private int _importPlacementEnumIndex { get; set; }
        private bool _manipulateAll { get; set; }
        public int importPlacementEnumIndex { get { return _importPlacementEnumIndex; } }
        public bool manipulateAll { get { return _manipulateAll; } }
        public List<string> ImportPlacementEnums { get; }


        private static readonly LinkWindow _instance = new LinkWindow();
        public static LinkWindow Instance { get { return _instance; } }

        private LinkWindow()
        {
            InitializeComponent();

            ImportPlacementEnums = new List<string>()
            {
                "基点到基点",
                "原点到原点",
                "共享坐标"
            };
            combobox_Location.ItemsSource = ImportPlacementEnums;
            //combobox_Location.DisplayMemberPath = "displayName";
            combobox_Location.SelectedValuePath = combobox_Location.SelectedIndex.ToString();
            combobox_Location.SelectedIndex = 0;

            _importPlacementEnumIndex = 0;
            _manipulateAll = false;
        }


        private void RadioButton_CAD_checked(object sender, RoutedEventArgs e)
        {
            if (this.IsLoaded)
            {
                grid_Revit.Visibility = System.Windows.Visibility.Hidden;
                grid_CAD.Visibility = System.Windows.Visibility.Visible;
            }
        }

        private void RadioButton_Revit_checked(object sender, RoutedEventArgs e)
        {
            if (this.IsLoaded)
            {
                grid_CAD.Visibility = System.Windows.Visibility.Hidden;
                grid_Revit.Visibility = System.Windows.Visibility.Visible;
            }
        }

        private void combobox_Location_selectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (this.IsLoaded) _importPlacementEnumIndex = combobox_Location.SelectedIndex;
        }

        private void RadioButton_Current_checked(object sender, RoutedEventArgs e)
        {
            if (this.IsLoaded) _manipulateAll = false;
        }
        private void RadioButton_All_checked(object sender, RoutedEventArgs e)
        {
            if (this.IsLoaded) _manipulateAll = true;
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            //this.Hide();
        }

        private void Window_Load(object sender, RoutedEventArgs e)
        {
            //MessageBox.Show($"{Process.GetCurrentProcess().SessionId}\n{Process.GetCurrentProcess().ProcessName}");
        }
    }
}
