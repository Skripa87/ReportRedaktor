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
                if (persone.Name
                    .ToLower()
                    .Contains("егорова"))
                {
                    persone.Startday = TimeSpan.Parse("08:00:00");
                    persone.Endday = TimeSpan.Parse("17:00:00");
                }
                else
                {
                    persone.Startday = TimeSpan.Parse("09:00:00");
                    persone.Endday = TimeSpan.Parse("18:00:00");
                }
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
                              ? (enters.Max(e => e.Time) > persone.Endday)
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

        private void SetFormat<T>(IXLRange range, T value)
        {
            range.Merge();
            range.Style.Font.FontColor = XLColor.Black;
            range.Style.Font.FontSize = 10;
            range.Style.Font.FontName = "Arial Cyr";
            range.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            range.Style.Font.Bold = true;
            range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            range.Style
                 .Border
                 .TopBorder = XLBorderStyleValues.Thin;
            range.Style
                 .Border
                 .RightBorder = XLBorderStyleValues.Thin;
            range.Style
                 .Border
                 .LeftBorder = XLBorderStyleValues.Thin;
            range.Style
                 .Border
                 .BottomBorder = XLBorderStyleValues.Thin;
            range.SetValue(value);
        }

        private void SetFormat<T>(IXLCell cell, T value)
        {
            cell.Style
                .Border
                .TopBorder = XLBorderStyleValues.Thin;
            cell.Style
                .Border
                .BottomBorder = XLBorderStyleValues.Thin;
            cell.Style
                .Border
                .LeftBorder = XLBorderStyleValues.Thin;
            cell.Style
                .Border
                .RightBorder = XLBorderStyleValues.Thin;
            cell.Style
                .Font
                .FontName = "Calibri";
            cell.SetValue(value);
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
            SetFormat(range, "Сводная таблица");
            range.Style
                 .Font
                 .Underline = XLFontUnderlineValues.Single;
            worksheet.Cell("B3")
                     .Value = "Период";
            worksheet.Cell("B3")
                     .Style
                     .Font
                     .Italic = true;
            worksheet.Cell("B3")
                     .Style
                     .Alignment
                     .Horizontal = XLAlignmentHorizontalValues.Center;
            worksheet.Cell("C3")
                     .Value = start.Month
                                   .ToString();
            worksheet.Cell("C3")
                     .Style
                     .Alignment
                     .Horizontal = XLAlignmentHorizontalValues.Center;
            range = worksheet.Range("A5:A6");
            SetFormat(range, "Дата");
            range = worksheet.Range("B5:B6");
            SetFormat(range, "ФИО");
            range = worksheet.Range("C5:D5");
            SetFormat(range, "События");
            SetFormat(worksheet.Cell("C6"), "приход");
            SetFormat(worksheet.Cell("D6"), "уход");
            int rowNumber = 8;
            end = end.AddDays(1);
            while (!DateTime.Equals(start, end))
            {
                foreach (var item in persones)
                {
                    var currentRow = worksheet.Row(rowNumber);
                    SetFormat(currentRow.Cell(1), start.ToLongDateString());
                    SetFormat(currentRow.Cell(2), item.Name);
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
                        SetFormat(range, "Не зарегистрирован");
                    }
                    else
                    {
                        SetFormat(currentRow.Cell(3), enter);
                        SetFormat(currentRow.Cell(4), outer);
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
                SetFormat(range, "Журнал событий входа-выхода");
                worksheet.Cell("B3")
                         .SetValue("Период")
                         .Style
                         .Font
                         .Italic = true;
                worksheet.Cell("C3")
                         .SetValue(startPeriod.Month.ToString())
                         .Style
                         .Font
                         .Italic = true;
                SetFormat(worksheet.Range("A5:A6"), "Дата");
                SetFormat(worksheet.Range("B5:B6"), "ФИО");
                SetFormat(worksheet.Range("C5:D5"), "События");
                SetFormat(worksheet.Range("F5:G5"), "Время");
                SetFormat(worksheet.Range("H5:I5"), "Штрафы");
                SetFormat(worksheet.Cell("F6"), "Поздно");
                SetFormat(worksheet.Cell("G6"), "Рано");
                SetFormat(worksheet.Cell("H6"), "Поздно");
                SetFormat(worksheet.Cell("I6"), "Рано");
                SetFormat(worksheet.Cell("C6"), "приход");
                SetFormat(worksheet.Cell("D6"), "уход");
                var currentRow = 7;
                var start = startPeriod;
                var end = endPeriod.AddDays(1);
                bool hollyday = false;
                while (start != end)
                {
                    hollyday = (start.DayOfWeek == DayOfWeek.Sunday ||
                                start.DayOfWeek == DayOfWeek.Saturday) 
                             ? true
                             : false;
                    var row = worksheet.Row(currentRow);
                    SetFormat(row.Cell(1), start.ToLongDateString());
                    SetFormat(row.Cell(2), persone.Name);
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
                    //***************************************************************************//
                    if (string.IsNullOrEmpty(enter) && 
                        string.IsNullOrEmpty(outer) && !hollyday) 
                    {
                        SetFormat(worksheet.Range(currentRow, 3, currentRow, 4), "Не зарегистрирован");
                    }
                    else
                    {
                        SetFormat(row.Cell(3), enter);
                        SetFormat(row.Cell(4), outer);
                    }
                    //*************************************************************************************//
                    if (string.IsNullOrEmpty(enter) && !string.IsNullOrEmpty(outer) && !hollyday)
                    {
                        SetFormat(row.Cell(6),"Не отметился");
                        SetFormat(row.Cell(8),100);
                    }
                    else if (!string.IsNullOrEmpty(enter) && TimeSpan.Parse(enter) > persone.Startday && !hollyday)
                    {
                        SetFormat(row.Cell(6), (TimeSpan.Parse(enter) - persone.Startday).ToString());
                        if ((TimeSpan.Parse(enter) - persone.Startday).Minutes > 15)
                        {
                            SetFormat(row.Cell(8), 100);
                        }
                        else
                        {
                            SetFormat(row.Cell(8), 0);
                        }
                    }
                    else
                    {
                        SetFormat(row.Cell(6), "");
                        SetFormat(row.Cell(8), 0);
                    }
                    if (!string.IsNullOrEmpty(enter) && string.IsNullOrEmpty(outer) && !hollyday)
                    {
                        SetFormat(row.Cell(7), "Не отметился");
                        SetFormat(row.Cell(9), 150);
                    }
                    else if (!string.IsNullOrEmpty(outer) && TimeSpan.Parse(outer) < persone.Endday && !hollyday)
                    {
                        SetFormat(row.Cell(7), (persone.Endday - TimeSpan.Parse(outer)).ToString());
                        SetFormat(row.Cell(9),150);
                    }
                    else 
                    {
                        SetFormat(row.Cell(7), "");
                        SetFormat(row.Cell(9), 0);
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
