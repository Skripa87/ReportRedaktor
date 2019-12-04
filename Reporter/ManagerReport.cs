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

        private List<Person> GetPersons()
        {
            var persons = new List<Person>();
            var workEvents = GetWorkEvents() ?? new List<WorkEvent>();
            var count = workEvents.Count;
            while (workEvents.Count > 0)
            {
                var person = new Person(workEvents.FirstOrDefault()
                                                  ?.UserName);
                person.SetWorkTime(FiveClockWorkers);
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
                    var eventsCurrentDateEnter = personEvents.FindAll(e => DateTime.Equals(e.Date, currentDate)
                                                                       && e.Direction == Direction.In);
                    var description = "";
                    var eventsCurrentDateOut = personEvents.FindAll(e =>
                                                   (DateTime.Equals(e.Date, currentDate) ||
                                                   (DateTime.Equals(e.Date, currentDate.AddDays(1))
                                                    && (e.Time > TimeSpan.Zero && e.Time < person.Startday)))
                                                    && e.Direction == Direction.Out);
                    var enter = TimeSpan.MinValue;
                    if (eventsCurrentDateEnter.Count == 0)
                    {
                        description = eventsCurrentDateOut.Count == 0
                                    ? "Административный"
                                    : (eventsCurrentDateOut.Count == 1
                                        ? "Не отметился, не возможно определить время прихода"
                                        : $"Не отметился, предполагаемое время прихода {TimeSpan.FromMilliseconds(eventsCurrentDateOut.ElementAt(0).Time.TotalMilliseconds - (eventsCurrentDateOut.ElementAt(1).Time.TotalMilliseconds / 2))}"
                                    );
                    }
                    else
                    {
                        enter = eventsCurrentDateEnter.FirstOrDefault()
                                                     ?.Time
                              ?? TimeSpan.MinValue;
                    }
                    var outer = TimeSpan.MinValue;
                    if (eventsCurrentDateOut.Count == 0)
                    {
                        description = eventsCurrentDateEnter.Count == 0
                                    ? "Административный"
                                    : (eventsCurrentDateEnter.Count == 1
                                        ? "Не отметился, не возможно определить время ухода"
                                        : $"Не отметился, последний зарегестрированный вход {eventsCurrentDateEnter.Last().Time}");
                    }
                    else
                    {
                        var index = personEvents.IndexOf(eventsCurrentDateOut.Last());
                        if (index == personEvents.Count - 1)
                        {
                            outer = eventsCurrentDateOut.Last()
                                                        .Time;
                        }
                        else
                        {
                            outer = personEvents.ElementAt(index + 1)
                                                .Direction == Direction.In
                                ? eventsCurrentDateOut.Last()
                                                      .Time
                                : personEvents.FindAll(p => DateTime.Equals(p.Date, currentDate) ||
                                                            (DateTime.Equals(p.Date, currentDate.AddDays(1)) &&
                                                             (p.Time < person.Startday)) &&
                                                            p.Direction == Direction.Out).Last()?.Time ?? TimeSpan.Zero;
                        }
                    }
                    person.VisitList
                          .Add(new Visit(currentDate, enter, outer, description));
                    personEvents.RemoveAll(p => DateTime.Equals(p.Date, currentDate));
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
            persones = GetPersons();
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
            var fifteenMinutes = TimeSpan.Parse("00:15:00");
            var hour = TimeSpan.Parse("01:00:00");
            progress.Value = 75;
            foreach (var person in persons)
            {
                var worksheet = workbook.AddWorksheet(person.Name.Substring(0, person.Name
                                                                                                   .Length >= 31
                                                                                             ? 31
                                                                                             : person.Name
                                                                                                     .Length));
                var countAllEvents = 0;
                worksheet = CreateHeaderWorkSheet(worksheet, startPeriod);
                var currentRow = 7;
                var start = startPeriod;
                var end = endPeriod.AddDays(1);
                var payForLateness = 0;
                var payForEarly = 0;
                TimeSpan latenessSum = TimeSpan.Zero, latenessEnterSum = TimeSpan.Zero, overWorkSum = TimeSpan.Zero, lateness = TimeSpan.Zero, earlyOut = TimeSpan.Zero, overWork = TimeSpan.Zero;
                int fullQuantityEnter = 0;
                var holiday = false;
                while (start != end)
                {
                    var row = worksheet.Row(currentRow);
                    if (start.DayOfWeek == DayOfWeek.Sunday || start.DayOfWeek == DayOfWeek.Saturday)
                    {
                        row.Style
                           .Fill
                           .BackgroundColor = XLColor.LightPink;
                        holiday = true;
                    }
                    else
                    {
                        holiday = false;
                    }
                    SetFormat(row.Cell(1), start.ToLongDateString());
                    SetFormat(row.Cell(2), person.Name);
                    var enter = person.VisitList
                                      .Find(v => DateTime.Equals(start, v.Date))
                                      ?.Enter ?? TimeSpan.MinValue;
                    var outer = person.VisitList
                                      .Find(v => DateTime.Equals(start, v.Date))
                                      ?.Outer ?? TimeSpan.MinValue;
                    var description = person.VisitList
                                            .Find(v => DateTime.Equals(start, v.Date))
                                            ?.Description;
                    //***************************************************************************//
                    if (TimeSpan.Equals(enter, TimeSpan.MinValue) && TimeSpan.Equals(outer, TimeSpan.MinValue))
                    {
                        if (!holiday)
                        {
                            SetFormat(worksheet.Range(currentRow, 3, currentRow, 4), "Не зарегистрирован");
                        }
                    }
                    else if (!TimeSpan.Equals(enter, TimeSpan.MinValue) && TimeSpan.Equals(outer, TimeSpan.MinValue))
                    {
                        SetFormat(row.Cell(3), $"{enter:hh}:{enter:mm}:{enter:ss}");
                        SetFormat(row.Cell(4), "");
                        SetFormat(row.Cell(5), description ?? "");
                        SetFormat(row.Cell(7), "Не отметился");
                        if (!holiday)
                        {
                            SetFormat(row.Cell(9), "150");
                            payForEarly += 150;
                            if (enter > person.Startday
                                    .Add(fifteenMinutes))
                            {
                                lateness = enter - person.Startday;
                                SetFormat(row.Cell(6), $"{lateness:hh}:{lateness:mm}:{lateness:ss}");
                                payForLateness += 100;
                                if (enter < person.Startday.Add(hour))
                                {
                                    SetFormat(row.Cell(8), "100");
                                }
                                else
                                {
                                    SetFormat(row.Cell(8), "?");
                                    row.Cell(8).Style.Fill.BackgroundColor = XLColor.Red;
                                }
                            }
                            else
                            {
                                SetFormat(row.Cell(6), "00:00");
                                SetFormat(row.Cell(8), "0");
                            }
                        }
                        else
                        {
                            SetFormat(row.Cell(9), "0");
                            SetFormat(row.Cell(6), "00:00");
                            SetFormat(row.Cell(8), "0");
                        }
                    }
                    else if(TimeSpan.Equals(enter, TimeSpan.MinValue) && !TimeSpan.Equals(outer, TimeSpan.MinValue))
                    {
                        SetFormat(row.Cell(3), "");
                        SetFormat(row.Cell(4), $"{outer:hh}:{outer:mm}:{outer:ss}");
                        SetFormat(row.Cell(5), description ?? "");
                        SetFormat(row.Cell(6), "Не отметился");
                        if (outer > person.Endday.Add(hour) && CheckWorkCount)
                        {
                            if (outer < person.Startday)
                            {
                                overWork = TimeSpan.FromDays(1)
                                    .Subtract(person.Endday
                                        .Add(hour))
                                    .Add(outer)
                                    .Subtract(hour)
                                    .Subtract(fifteenMinutes);
                                overWork = TimeSpan.FromMinutes(overWork.Minutes - (overWork.Minutes % 30));
                            }
                            else
                            {
                                overWork = outer - person.Endday.Add(hour).Add(fifteenMinutes);
                                overWork = TimeSpan.FromMinutes(overWork.Minutes - (overWork.Minutes % 30));
                            }
                            SetFormat(row.Cell(11), $"{overWork:hh}:{overWork:mm}:{overWork:ss}");
                            overWorkSum += overWork;
                        }
                        else if (!holiday)
                        {
                            SetFormat(row.Cell(8), "100");
                            payForLateness += 100;
                            if (outer < person.Endday && outer > person.Startday)
                            {
                                earlyOut = person.Endday - -outer;
                                SetFormat(row.Cell(7), $"{earlyOut:hh}:{earlyOut:mm}:{earlyOut:ss}");
                                payForEarly += 100;
                                if (outer > person.Endday.Subtract(hour))
                                {
                                    SetFormat(row.Cell(9), "100");
                                }
                                else
                                {
                                    SetFormat(row.Cell(9), "?");
                                    row.Cell(9).Style.Fill.BackgroundColor = XLColor.Red;
                                }
                            }
                            else
                            {
                                SetFormat(row.Cell(7), "00:00");
                                SetFormat(row.Cell(9), "0");
                            }
                        }
                        else
                        {
                            SetFormat(row.Cell(8), "0");
                            SetFormat(row.Cell(7), "00:00");
                            SetFormat(row.Cell(9), "0");
                        }
                    }else if (!TimeSpan.Equals(enter, TimeSpan.MinValue) && !TimeSpan.Equals(outer, TimeSpan.MinValue))
                    {
                        SetFormat(row.Cell(3), $"{enter:hh}:{enter:mm}:{enter:ss}");
                        SetFormat(row.Cell(4), $"{outer:hh}:{outer:mm}:{outer:ss}");
                        SetFormat(row.Cell(5), description ?? "");
                        if (holiday)
                        {
                            if (outer < enter)
                            {
                                overWork = TimeSpan.FromDays(1).Subtract(person.Endday.Add(hour)).Add(outer);
                            }
                        }
                        else 
                        {
                            if (enter > person.Startday
                                              .Add(fifteenMinutes))
                            {
                                lateness = enter - person.Startday
                                                         .Add(fifteenMinutes);
                                latenessEnterSum += lateness;
                                SetFormat(row.Cell(6), lateness);
                                if (lateness > hour)
                                {
                                    SetFormat(row.Cell(8), "?");
                                    row.Cell(8)
                                       .Style
                                       .Fill
                                       .BackgroundColor = XLColor.Red;
                                }
                                else
                                {
                                    SetFormat(row.Cell(8),"100");
                                    payForLateness += 100;
                                }
                            }

                            if (outer < person.Endday && outer > person.Startday)
                            {
                                earlyOut = person.Endday - outer;
                                if (earlyOut > hour)
                                {
                                    SetFormat(row.Cell(9), "?");
                                    row.Cell(9)
                                        .Style
                                        .Fill
                                        .BackgroundColor = XLColor.Red;
                                }
                                else
                                {
                                    SetFormat(row.Cell(9),150);
                                    payForEarly += 150;
                                }
                            }
                        }
                    }
                    start = start.AddDays(1);
                }
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
