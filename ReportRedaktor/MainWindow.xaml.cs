using System;
using System.Collections.Generic;
using System.Linq;
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
using DocumentFormat.OpenXml.Office2013.Excel;

namespace Reporter
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string fileNameExcel;        

        public MainWindow()
        {
            InitializeComponent();
            checkBox.IsChecked = false;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(fileNameExcel))
                return;
            var report_Manager = new ManagerReport(fileNameExcel,checkBox.IsChecked ?? false);
            DateTime start = (DateTime)(calendarStart.SelectedDate ?? DateTime.MinValue);
            DateTime end = (DateTime)(calendarEnd.SelectedDate ?? DateTime.MaxValue);
            //var end = DateTime.Parse("31.08.2019");
            report_Manager.GetReport(start,end, progressBar);
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.FileName = "Document";
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Text documents (.xlsx)|*.xlsx";

            Nullable<bool> result = dlg.ShowDialog();

            if (result == true)
            {
                fileNameExcel = dlg.FileName;
            }
        }

        private void ProgressBar_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {

        }        
    }
}
