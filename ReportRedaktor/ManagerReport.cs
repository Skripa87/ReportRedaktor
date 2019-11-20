using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Windows.Controls;
using OfficeOpenXml;
using ClosedXML.Excel;

namespace Reporter
{
    public class ManagerReport
    {
        private static List<string> FiveClockWorkers { get; set; }
        private static bool CheckWorkCount { get; set; }
        public ManagerReport(string fileName, bool checkWorkCount)
        {
            FileName = fileName;
            CheckWorkCount = checkWorkCount;
            FiveClockWorkers = new List<string> { "камалетдинов а", "трофимов а", "галиуллин в", "егорова ю", "тазетдинов р", "тазетдинов а", "шангареева г" };
        }

        private string FileName { get; set; }

        private List<WorkEvent> GetWorkEvents()
        {
            var workEvents = new List<WorkEvent>();
            string fileNameExcel = FileName;
            var failInfo = new FileInfo(fileNameExcel);
            using (var package = new ExcelPackage(failInfo))
            {
                var epWorkbook = package.Workbook;
                var worksheet = epWorkbook.Worksheets
                                            .First();
                var end = worksheet.Dimension.End;
                for (int row = 9; row < end.Row; row++)
                {
                    var workEvent = new WorkEvent(worksheet.GetValue(row, 1).ToString(),
                                                   worksheet.GetValue(row, 2).ToString(),
                                                   worksheet.GetValue(row, 3).ToString(),
                                                   worksheet.GetValue(row, 4).ToString(),
                                                   worksheet.GetValue(row, 5)?.ToString(),
                                                   "",
                                                   worksheet.GetValue(row, 8)?.ToString() ?? "");
                    workEvents.Add(workEvent);
                }
            }
            return workEvents.FindAll(f => !string.IsNullOrEmpty(f.UserName));
        }

        private List<Person> GetPersons(ProgressBar progress)
        {
            var persons = new List<Person>();
            var workEvents = GetWorkEvents() ?? new List<WorkEvent>();
            var count = workEvents.Count;
            while (workEvents.Count > 0)
            {
                var progressbar = (count - workEvents.Count) * 25 / count;
                progress.Value = progressbar;
                progress.UpdateLayout();
                var person = new Person(workEvents.FirstOrDefault()
                                                  ?.UserName);
                if (FiveClockWorkers.Any(f => person.Name
                                                     .ToLower().Contains(f)))
                {
                    person.Startday = TimeSpan.Parse("08:00:00");
                    person.Endday = TimeSpan.Parse("17:00:00");
                }
                else
                {
                    person.Startday = TimeSpan.Parse("09:00:00");
                    person.Endday = TimeSpan.Parse("18:00:00");
                }
                var personEvents = workEvents.FindAll(e => string.Equals(e.UserName, person.Name));
                workEvents.RemoveAll(e => string.Equals(e.UserName, person.Name));
                while (personEvents.Count > 0)
                {
                    WorkEvent first = null;
                    foreach (var @event in personEvents)
                    {
                        first = @event;
                        break;
                    }
                    if (first == null) continue;
                    var currentDate = first.Date;
                    var eventsCurrentDateEnter = personEvents.FindAll(e =>e.Date>currentDate 
                                                                       && e.Direction == Direction.In);
                    if (eventsCurrentDateEnter.Count == 0)
                    {

                    }
                    var eventsCurrentDateOut = personEvents.FindAll(e =>
                                                   DateTime.Equals(e.Date, currentDate) 
                                                                && e.Direction == Direction.Out);
                    //var visit = new Visit(currentDate, enter, outer);
                    person.VisitList.Add(new Visit(currentDate,));
                }
                persons.Add(person);
            }
            return persons;
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

        private XLWorkbook GetReportForPeriod(DateTime start, DateTime end, out List<Person> persones, ProgressBar progress)
        {
            persones = GetPersons(progress);
            persones.RemoveAll(p => p.Name
                                        .ToLower()
                                        .Contains("гость") ||
                                    string.IsNullOrEmpty(p.Name));
            persones.Sort();
            var newFileName = FileName.Substring(0, FileName.Length - 5) + "_new.xlsx";
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
                    var enter = buffer != null && buffer != TimeSpan.Zero
                              ? buffer.ToString()
                              : "";
                    buffer = item.VisitList
                                 .Find(v => DateTime.Equals(v.Date, start))
                                ?.Outer;
                    var outer = buffer != null && buffer != TimeSpan.Zero
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

        private IXLWorksheet CreateHeaderWorkSheet(IXLWorksheet worksheet, DateTime startPeriod)
        {
            var range = worksheet.Range("A1:D1");
            SetFormat(range, "Журнал событий входа-выхода");
            worksheet.Cell("B3")
                     .SetValue("Период")
                     .Style
                     .Font
                     .Italic = true;
            worksheet.Cell("C3")
                     .SetValue(startPeriod.ToString("MMMM"))
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
            if (CheckWorkCount)
            {
                SetFormat(worksheet.Range("K5:K6"), "Переработка");
            }
            return worksheet;
        }

        private void AddPersonalReports(XLWorkbook workbook, List<Person> persons, DateTime startPeriod,
            DateTime endPeriod, ProgressBar progress)
        {
            progress.Value = 75;
            foreach (var person in persons)
            {
                var worksheet = workbook.AddWorksheet(person.Name.Substring(0, person.Name
                                                                                                   .Length >= 31
                                                                                             ? 31
                                                                                             : person.Name
                                                                                                     .Length));
                var countAllEvents = 0;
                CreateHeaderWorkSheet(worksheet, startPeriod);
                var currentRow = 7;
                var start = startPeriod;
                var end = endPeriod.AddDays(1);
                TimeSpan latenessSum = TimeSpan.Zero, latenessEnterSum = TimeSpan.Zero, overWorkSum = TimeSpan.Zero;
                int fullQuantityEnter = 0;
                while (start != end)
                {
                    var row = worksheet.Row(currentRow);
                    if (start.DayOfWeek == DayOfWeek.Sunday || start.DayOfWeek == DayOfWeek.Saturday)
                    {
                        row.Style.Fill.BackgroundColor = XLColor.LightPink;




                    }

                    
                    
                            
                    //        SetFormat(row.Cell(1), start.ToLongDateString());
                    //        SetFormat(row.Cell(2), person.Name);
                    //        var enterTimeSpan = person.VisitList
                    //                           .Find(v => DateTime.Equals(start, v.Date))
                    //                           ?.Enter ?? TimeSpan.Zero;
                    //        var enter = enterTimeSpan == TimeSpan.Zero
                    //                  ? ""
                    //                  : enterTimeSpan.ToString();
                    //        var outerTimeSpan = person.VisitList
                    //                       .Find(v => DateTime.Equals(start, v.Date))
                    //                       ?.Outer ?? TimeSpan.Zero;
                    //        var outer = outerTimeSpan == TimeSpan.Zero
                    //                  ? ""
                    //                  : outerTimeSpan.ToString();
                    //        //***************************************************************************//
                    //        if (string.IsNullOrEmpty(enter) && 
                    //            string.IsNullOrEmpty(outer) && !holiday) 
                    //        {
                    //            SetFormat(worksheet.Range(currentRow, 3, currentRow, 4), "Не зарегистрирован");
                    //        }
                    //        else
                    //        {
                    //            SetFormat(row.Cell(3), enter);
                    //            SetFormat(row.Cell(4), outer);
                    //        }
                    //        //*************************************************************************************//
                    //        if (string.IsNullOrEmpty(enter) && !string.IsNullOrEmpty(outer) && !holiday)
                    //        {
                    //            SetFormat(row.Cell(6),"Не отметился");
                    //            SetFormat(row.Cell(8),100);
                    //            fullQuantityEnter += 100;
                    //        }
                    //        else if (!string.IsNullOrEmpty(enter) && enterTimeSpan > person.Startday && !holiday)
                    //        {
                    //            var latenessEnter =  enterTimeSpan - person.Startday;
                    //            SetFormat(row.Cell(6), $"{latenessEnter:hh}:{latenessEnter:mm}:{latenessEnter:ss}");
                    //            if (latenessEnter.Minutes > 15)
                    //            {
                    //                SetFormat(row.Cell(8), 100);
                    //                fullQuantityEnter += 100;
                    //            }
                    //            latenessEnterSum += latenessEnter;
                    //        }
                    //        else
                    //        {
                    //            SetFormat(row.Cell(6), "00:00");
                    //            SetFormat(row.Cell(8), 0);
                    //        }
                    //        if (!string.IsNullOrEmpty(enter) && string.IsNullOrEmpty(outer) && !holiday)
                    //        {
                    //            SetFormat(row.Cell(7), "Не отметился");
                    //            SetFormat(row.Cell(9), 150);
                    //            countAllEvents += 150;
                    //        }
                    //        else if (!string.IsNullOrEmpty(outer) && (outerTimeSpan < person.Endday 
                    //                                              && outerTimeSpan > TimeSpan.Parse("05:00:00")) 
                    //                                              && !holiday)
                    //        {
                    //            var lateness = person.Endday - outerTimeSpan;
                    //            SetFormat(row.Cell(7), $"{lateness:hh}:{lateness:mm}:{lateness:ss}");
                    //            SetFormat(row.Cell(9),150);
                    //            countAllEvents += 150;
                    //            latenessSum += lateness;
                    //        }
                    //        else 
                    //        {
                    //            SetFormat(row.Cell(7), "00:00");
                    //            SetFormat(row.Cell(9), 0);
                    //        }
                    //        TimeSpan overWork;
                    //        if (!string.IsNullOrEmpty(outer) && (outerTimeSpan > person.Endday
                    //                                             || outerTimeSpan < TimeSpan.Parse("05:00:00"))
                    //                                         && CheckWorkCount
                    //                                         && !holiday)
                    //        {
                    //            overWork = outerTimeSpan < TimeSpan.Parse("05:00:00") && outerTimeSpan != TimeSpan.Zero
                    //                     ? TimeSpan.Parse("06:00:00") + outerTimeSpan
                    //                     : outerTimeSpan - person.Endday;
                    //            SetFormat(row.Cell(11), $"{overWork:hh}:{overWork:mm}:{overWork:ss}");
                    //            overWorkSum += overWork;
                    //        }
                    //        else if(holiday && CheckWorkCount && !string.IsNullOrEmpty(outer))
                    //        {
                    //            overWork = outerTimeSpan < TimeSpan.Parse("05:00:00")
                    //                     ? TimeSpan.Parse("23:00:00") - enterTimeSpan + outerTimeSpan +
                    //                       TimeSpan.Parse("01:00:00")
                    //                     : outerTimeSpan - enterTimeSpan;
                    //            SetFormat(row.Cell(11), $"{overWork:hh}:{overWork:mm}:{overWork:ss}");
                    //            overWorkSum += overWork;
                    //        }
                    //        currentRow++;
                            start = start.AddDays(1);
                        }
                    //    currentRow++;
                    //    var summaryRow = worksheet.Row(currentRow);
                    //    SetFormat(summaryRow.Cell(6),latenessEnterSum);
                    //    SetFormat(summaryRow.Cell(7),latenessSum);
                    //    SetFormat(summaryRow.Cell(8),fullQuantityEnter);
                    //    SetFormat(summaryRow.Cell(9),countAllEvents);
                    //    if (CheckWorkCount)
                    //    {
                    //        SetFormat(summaryRow.Cell(11), overWorkSum);
                    //    }
                    //    var sumFullRange = worksheet.Range(currentRow + 1, 6, currentRow + 1, 7);
                    //    var fullRange = worksheet.Range(currentRow + 1, 8, currentRow + 1, 9);
                    //    sumFullRange.Merge();
                    //    fullRange.Merge();
                    //    SetFormat(sumFullRange,latenessSum + latenessEnterSum);
                    //    SetFormat(fullRange,countAllEvents + fullQuantityEnter);
                    //    worksheet.Columns().AdjustToContents();
                    //    worksheet.Rows().AdjustToContents();
                    //}

                    //try
                    //{
                    //    workbook.Save();
                    //}
                    //catch (Exception ex)
                    //{
                    //    // ignored
                    //}
                }
            }

            public void GetReport(DateTime start, DateTime end, ProgressBar progressBar)
            {
                var workBook = GetReportForPeriod(start, end, out var persons, progressBar);
                AddPersonalReports(workBook, persons, start, end, progressBar);
                progressBar.Value = 100;
            }
        }
    }
