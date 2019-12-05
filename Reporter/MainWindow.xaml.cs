using System;
using System.Collections.Generic;
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
            Holidays.Text = "";
            SelectedFile.Content = "";
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_fileNameExcel))
                return;
            var reportManager = new ManagerReport(_fileNameExcel, CheckBox.IsChecked ?? false, CheckBoxForCountWorker.IsChecked ?? false);
            DateTime start = CalendarStart.SelectedDate ?? DateTime.MinValue;
            DateTime end = CalendarEnd.SelectedDate ?? DateTime.MaxValue;
            //var end = DateTime.Parse("31.08.2019");
            char[] separator = new char[]{' ', ',',';','.','_'};
            var holidays = new List<int>();
            if (!string.IsNullOrEmpty(Holidays.Text))
            {
                var bufferHolidays = Holidays.Text
                                             .Split(separator);
                foreach (var item in bufferHolidays)
                {
                    if (int.TryParse(item, out var element))
                    {
                        if (!holidays.Contains(element))
                        {
                            holidays.Add(element);
                        }
                    }
                }
            }
            reportManager.GetReport(start, end, holidays, ProgressBar);
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
                SelectedFile.Content = _fileNameExcel;
            }
        }
        private void ProgressBar_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {

        }

        private void CheckBox_OnChecked(object sender, RoutedEventArgs e)
        {
            
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void Holidays_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {

        }
    }
}
