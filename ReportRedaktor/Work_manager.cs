using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Security.Cryptography;
using System.Windows.Controls;
using OfficeOpenXml;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;

namespace ReportRedaktor
{
    public class Report_manager
    {
        public Report_manager(string fileName)
        {
            File_name = fileName;
        }

        private string File_name { get; set; }

        private List<Work_event> GetWork_Events()
        {
            var work_events = new List<Work_event>();
            string fileNameExcel = File_name;
            var failInfo = new FileInfo(fileNameExcel);
            using (var package = new ExcelPackage(failInfo))
            {
                var ep_workbook = package.Workbook;
                var workscheet = ep_workbook.Worksheets
                                            .First();
                var start = workscheet.Dimension.Start;
                var end = workscheet.Dimension.End;
                for (int row = 9; row < end.Row; row++)
                {
                    var workEvent = new Work_event(workscheet.GetValue(row, 1).ToString(),
                                                   workscheet.GetValue(row, 2).ToString(),
                                                   workscheet.GetValue(row, 3).ToString(),
                                                   workscheet.GetValue(row, 4).ToString(),
                                                   workscheet.GetValue(row, 5)?.ToString(),
                                                   "",
                                                   workscheet.GetValue(row, 8)?.ToString() ?? "");
                    work_events.Add(workEvent);
                }
            }
            return work_events;
        }

        private List<Persone> GetPersones(ProgressBar progress)
        {
            var persones = new List<Persone>();
            var work_events = GetWork_Events();
            var count = work_events.Count;
            while (work_events.Count > 0)
            {
                var currentvalueprogressbar = (count - work_events.Count) * 25 / count;
                progress.Value = currentvalueprogressbar;
                progress.UpdateLayout();
                var persone = new Persone(work_events.FirstOrDefault().UserName);
                var personeEvents = work_events.FindAll(e => string.Equals(e.UserName, persone.Name));
                work_events.RemoveAll(e => string.Equals(e.UserName, persone.Name));
                while (personeEvents.Count > 0)
                {
                    var curent_date = personeEvents.FirstOrDefault().Date;
                    var events = personeEvents.FindAll(p => DateTime.Equals(p.Date, curent_date));
                    personeEvents.RemoveAll(p => DateTime.Equals(p.Date, curent_date));
                    var enters = events.FindAll(e => e.Direction == Direction.IN);
                    var enter = enters.Count > 0
                              ? enters.Min(e => e.Time)
                              : TimeSpan.MinValue;
                    var outers = events.FindAll(e => e.Direction == Direction.OUT);
                    var outer = outers.Count == 0 ||
                               (enters.Count > 0 &&
                                enters.Any(e => e.Time > outers.Max(o => o.Time)))
                              ? (enters.Max(e => e.Time) > TimeSpan.Parse("17:00:00"))
                                ? enters.Max(e => e.Time)
                                : TimeSpan.MinValue
                              : outers.Max(e => e.Time);
                    var visit = new Visit(curent_date, enter, outer);
                    persone.VisitList.Add(visit);
                }
                persones.Add(persone);
            }
            return persones;
        }

        private XLWorkbook GetReportForPeriod(DateTime start, DateTime end, out List<Persone> persones, ProgressBar progress)
        {
            persones = GetPersones(progress);
            persones.RemoveAll(p => p.Name
                                        .ToLower()
                                        .Contains("гость") ||
                                    string.IsNullOrEmpty(p.Name));
            persones.Sort();
            var newFileName = File_name.Substring(0, File_name.Length - 5) + "_new.xlsx";
            XLWorkbook workbook = new XLWorkbook();
            var worksheet = workbook.AddWorksheet("Сводная_таблица");
            var range = worksheet.Range("A1:D1");
            range.Merge()
                 .Value = "Сводная таблица";
            range.Style
                 .Font
                 .Bold = true;
            range.Style.Font.FontColor = XLColor.Black;
            range.Style.Font.FontName = "Arial Cyr";
            range.Style.Font.Underline = XLFontUnderlineValues.Single;
            worksheet.Cell("B3")
                     .Value = "Период";
            worksheet.Cell("B3").Style
                                .Font
                                .Italic = true; 
            worksheet.Cell("C3")
                     .Value = start.Month
                                   .ToString();
            range = worksheet.Range("A5:A6");
            range.Merge()
                 .SetValue<string>("Дата");
            range.Style.Font.FontSize = 10;
            range.Style.Font.FontName = "Arial Cyr";
            range.Style.Font.SetVerticalAlignment(XLFontVerticalTextAlignmentValues.Baseline);
            range.Style.Font.Bold = true;
            range = worksheet.Range("B5:B6");
            range.Merge()
                 .SetValue<string>("ФИО");
            range.Style.Font.FontSize = 10;
            range.Style.Font.FontName = "Arial Cyr";
            range.Style.Font.SetVerticalAlignment(XLFontVerticalTextAlignmentValues.Baseline);
            range.Style.Font.Bold = true;
            range = worksheet.Range("C5:D5");
            range.Merge()
                 .SetValue<string>("События");
            worksheet.Cell("C6")
                     .Value = "приход";
            worksheet.Cell("D6")
                     .Value = "уход";
            int rowNumber = 8;
            end = end.AddDays(1);
            while (!DateTime.Equals(start, end))
            {
                foreach (var item in persones)
                {
                    var currentRow = worksheet.Row(rowNumber);
                    currentRow.Cell(1)
                              .SetValue<DateTime>(start);
                    currentRow.Cell(2)
                              .SetValue(item.Name);
                    var buffer = item.VisitList
                                     .Find(v => DateTime.Equals(v.Date, start))
                                    ?.Enter;
                    var enter = buffer != null && buffer != TimeSpan.MinValue
                              ? buffer.ToString()
                              : "";
                    buffer = item.VisitList
                                 .Find(v => DateTime.Equals(v.Date, start))
                                ?.Outer;
                    var outer = buffer != null && buffer != TimeSpan.MinValue
                              ? buffer.ToString()
                              : "";
                    if (string.IsNullOrEmpty(enter) && string.IsNullOrEmpty(outer))
                    {
                        range = worksheet.Range(rowNumber, 3, rowNumber, 4);
                        range.Merge()
                             .SetValue("Не зарегистрирован");
                    }
                    else
                    {
                        currentRow.Cell(3)
                                  .SetValue(enter);
                        currentRow.Cell(4)
                                  .SetValue(outer);
                    }
                    rowNumber++;
                }
                start = start.AddDays(1);
            }
            worksheet.Columns().AdjustToContents();
            worksheet.Rows().AdjustToContents();
            workbook.SaveAs(newFileName);
            progress.Value = 50;
            progress.UpdateLayout();
            return workbook;
        }

        private void AddPersonalReports(XLWorkbook workbook, List<Persone> persones, DateTime startPeriod, DateTime endPeriod, ProgressBar progress)
        {
            progress.Value = 75;
            foreach (var persone in persones)
            {
                var worksheet = workbook.AddWorksheet(persone.Name.Substring(0, persone.Name
                                                                                      .Length >= 31
                                                                              ? 31
                                                                              : persone.Name
                                                                                       .Length));
                var range = worksheet.Range("A1:D1");
                range.Merge()
                     .SetValue("Журнал событий входа-выхода");
                worksheet.Cell("B3")
                         .SetValue("Период");
                worksheet.Cell("C3")
                         .SetValue(startPeriod.Month.ToString());
                range = worksheet.Range("A5:A6");
                range.Merge()
                     .SetValue("Дата");
                range = worksheet.Range("B5:B6");
                range.Merge()
                     .SetValue("ФИО");
                range = worksheet.Range("C5:D5");
                range.Merge()
                     .SetValue("События");
                range = worksheet.Range("F5:G5");
                range.Merge().SetValue("Время");
                range = worksheet.Range("H5:I5");
                range.Merge().SetValue("Штрафы");
                worksheet.Cell("F6")
                         .SetValue("Поздно");
                worksheet.Cell("G6")
                         .SetValue("Рано");
                worksheet.Cell("H6")
                         .SetValue("Поздно");
                worksheet.Cell("I6")
                         .SetValue("Рано");
                worksheet.Cell("C6")
                         .SetValue("приход");
                worksheet.Cell("D6")
                         .SetValue("уход");
                var currentRow = 7;
                var start = startPeriod;
                var end = endPeriod.AddDays(1);
                while (start != end)
                {
                    var row = worksheet.Row(currentRow);
                    row.Cell(1)
                       .SetValue(start.ToLongDateString());
                    row.Cell(2)
                       .SetValue(persone.Name);
                    var bufer = persone.VisitList
                                       .Find(v => DateTime.Equals(start, v.Date))
                                      ?.Enter;
                    var enter = bufer == null || bufer == TimeSpan.MinValue
                              ? ""
                              : bufer.ToString();
                    bufer = persone.VisitList
                                   .Find(v => DateTime.Equals(start, v.Date))
                                   ?.Outer;
                    var outer = bufer == null || bufer == TimeSpan.MinValue
                              ? ""
                              : bufer.ToString();
                    if (string.IsNullOrEmpty(enter) && string.IsNullOrEmpty(outer))
                    {
                        range = worksheet.Range(currentRow, 3, currentRow, 4);
                        range.Merge()
                             .SetValue("Не зарегистрирован");
                    }
                    else
                    {
                        row.Cell(3).SetValue(enter);
                        row.Cell(4).SetValue(outer);
                    }
                    if (outer == "")
                    {
                        row.Cell(7)
                           .SetValue("");
                    }
                    else
                        if (TimeSpan.Parse(outer) < TimeSpan.Parse("18:00:00"))
                    {
                        row.Cell(7)
                           .SetValue((TimeSpan.Parse("18:00:00") - TimeSpan.Parse(outer))
                                                                           .ToString());
                    }
                    if (enter == "")
                    {
                        row.Cell(7)
                              .SetValue("");
                    }
                    else 
                    if(TimeSpan.Parse(enter) > TimeSpan.Parse("9:00:00"))
                    {
                        row.Cell(6)
                           .SetValue((TimeSpan.Parse(enter) - TimeSpan.Parse("9:00:00"))
                                                                      .ToString());
                    }
                    currentRow++;
                    start = start.AddDays(1);
                }
                worksheet.Columns().AdjustToContents();
                worksheet.Rows().AdjustToContents();
            }
            workbook.Save();
        }

        public void GetReport(DateTime start, DateTime end, ProgressBar progressBar)
        {
            var workBook = GetReportForPeriod(start, end, out var persones, progressBar);
            AddPersonalReports(workBook, persones, start, end, progressBar);
            progressBar.Value = 100;
        }
    }
}
