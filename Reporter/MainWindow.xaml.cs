using System;
using System.Windows;

namespace Reporter
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private string _fileNameExcel;        

        public MainWindow()
        {
            InitializeComponent();
            CheckBox.IsChecked = false;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_fileNameExcel))
                return;
            var reportManager = new ManagerReport(_fileNameExcel,CheckBox.IsChecked ?? false);
            DateTime start = CalendarStart.SelectedDate ?? DateTime.MinValue;
            DateTime end = CalendarEnd.SelectedDate ?? DateTime.MaxValue;
            //var end = DateTime.Parse("31.08.2019");
            reportManager.GetReport(start,end, ProgressBar);
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
                _fileNameExcel = dlg.FileName;
            }
        }

        private void ProgressBar_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {

        }        
    }
}
