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
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2013.Excel;

namespace ReportRedaktor
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string fileNameExcel;

        private List<Day> getPeriod(string fileName)
        {
            fileNameExcel = fileName;
            var workbook = new XLWorkbook(fileNameExcel);
            var workcheet = workbook.Worksheet(1);
            int row = 9, column = 1;
            string date = "0", name = "", enter = "", _out = "";
            List<Day> Period = new List<Day>();
            Day day = new Day(date);
            Persone persone = new Persone("");
            while (true)
            {
                date = workcheet.Cell(row, column)
                                .Value.ToString();
                if (date != "" && DateTime.Parse(date)
                                          .ToString() != day.Date.ToString())
                {
                    if (day.Date != null && day.Persones.Count > 0)
                    {
                        Period.Add(day);
                        day = new Day(date);
                    }
                }
                column += 2;
                name = workcheet.Cell(row, column)
                                .Value
                                .ToString();
                if (name != "" && persone.Name != name)
                {
                    day.Persones.Add(persone);
                    persone = new Persone(name);
                }
                column += 2;
                enter = workcheet.Cell(row, column)
                                .Value
                                .ToString();
                persone.EnteredList.Add(enter);
                column++;
                _out = workcheet.Cell(row, column)
                                .Value
                                .ToString();
                persone.GetOutList
                       .Add(_out);
                column = 1;
                row++;
                if (date == "" && name == "" && enter == "" && _out == "")
                {
                    day.Persones.Add(persone);
                    Period.Add(day);
                    break;
                }
            }
            return Period;
        }

        private void createPersonalReports(XLWorkbook report, List<Day> period)
        {
            List<string> personNames = new List<string>();
            foreach (var item in period)
            {
                var current_names = item.Persones
                                        .Select(s => s.Name)
                                        .ToList();
                var bufferPersoneNames = personNames;
                personNames.AddRange(current_names.Except(bufferPersoneNames.Intersect(current_names)));
            }
            List<PersonalReport> personalReports = new List<PersonalReport>();
            foreach (var item in personNames)
            {
                PersonalReport personalReport = new PersonalReport() {Name = item };
                foreach (var day in period)
                {
                    var visit = new Visit()
                    {
                        Date = day.Date.ToShortDateString(),
                        Enter = day.Persones
                                   .Find(p => string.Equals(p.Name, personalReport.Name))
                                  ?.EnteredList
                                  ?.FirstOrDefault(),
                        Outer = day.Persones
                                   .Find(p => string.Equals(p.Name, personalReport.Name))
                                  ?.GetOutList
                                  ?.LastOrDefault()
                    };
                    personalReport.Visits
                                  .Add(visit);
                }
                personalReports.Add(personalReport);
            }
            foreach (var item in personalReports)
            {
                var ws = report.Worksheets.Add(item.Name);
                int row = 1, column = 1;
                foreach (var visit in item.Visits)
                {
                    ws.Cell(row, column).SetValue<string>(visit.Date);
                    column++;
                    ws.Cell(row, column).SetValue<string>(item.Name);
                    column++;
                    ws.Cell(row, column).SetValue<string>(visit.Enter);
                    column++;
                    ws.Cell(row, column).SetValue<string>(visit.Outer);
                    row++;
                    column = 1;
                }
                ws.Columns().AdjustToContents();
            }
        }

        private void getValidReport(List<Day> period, string fileNameReport)
        {
            var report = new XLWorkbook();
            var ws = report.Worksheets.Add("Свод");
            int row = 1, column = 1;
            string date = "";
            period.Remove(period.FirstOrDefault());
            foreach (var item in period)
            {
                date = item.Date.ToLongDateString();
                foreach (var persone in item.Persones)
                {
                    ws.Cell(row, column).SetValue<string>(date);
                    column += 2;
                    ws.Cell(row, column).SetValue<string>(persone.Name);
                    column++;
                    ws.Cell(row, column).SetValue<string>(persone.EnteredList
                                                                 .FirstOrDefault());
                    column++;
                    ws.Cell(row, column).SetValue<string>(persone.GetOutList
                                                                 .LastOrDefault());
                    column = 1;
                    row++;
                }
            }
            createPersonalReports(report, period);
            report.SaveAs(fileNameReport.Substring(0,fileNameReport.IndexOf('.') - 1) + "_report" + ".xlsx");
        }

        public MainWindow()
        {
            InitializeComponent();            
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(fileNameExcel))
                return;
            var work_manager = new Work_manager();
            var result = work_manager.GetWork_Events(fileNameExcel);
            //List<Day> Period = getPeriod(fileNameExcel);
            //getValidReport(Period, fileNameExcel);
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.FileName = "Document"; // Default file name
            dlg.DefaultExt = ".xlsx"; // Default file extension
            dlg.Filter = "Text documents (.xlsx)|*.xlsx"; // Filter files by extension

            // Show open file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                // Open document
                fileNameExcel = dlg.FileName;
            }
        }
    }
}
