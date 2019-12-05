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

        private static bool CountWorker { get; set; }
        private static TimeSpan FifteenMinuts { get; set; }
        private static TimeSpan Hour { get; set; }
        private static TimeSpan Day { get; set; }
        public ManagerReport(string fileName, bool checkWorkCount, bool countWorker)
        {
            FileName = fileName;
            CheckWorkCount = checkWorkCount;
            CountWorker = countWorker;
            FiveClockWorkers = new List<string> { "камалетдинов а", "трофимов а", "галиуллин в", "егорова ю", "тазетдинов р", "тазетдинов а", "шангареева г" };
            FifteenMinuts = new TimeSpan(0,0,15,0);
            Hour = new TimeSpan(0,1,0,0);
            Day = new TimeSpan(1,0,0,0);
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

        private XLWorkbook GetReportForPeriod(DateTime start, DateTime end, out List<Person> persons, ProgressBar progress)
        {
            persons = GetPersons();
            persons.RemoveAll(p => p.Name
                                         .ToLower()
                                         .Contains("гость") ||
                                    string.IsNullOrEmpty(p.Name));
            persons.Sort();
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
                foreach (var item in persons)
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

        /*Создание шапки для персонального отчета сотрудника*/
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
            SetFormat(worksheet.Range("A5:A6"), @"Дата");
            SetFormat(worksheet.Range("B5:B6"), @"ФИО");
            SetFormat(worksheet.Range("C5:D5"), @"События");
            SetFormat(worksheet.Range("F5:G5"), @"Время");
            SetFormat(worksheet.Range("H5:I5"), @"Штрафы");
            SetFormat(worksheet.Cell("F6"), @"Поздно");
            SetFormat(worksheet.Cell("G6"), @"Рано");
            SetFormat(worksheet.Cell("H6"), @"Поздно");
            SetFormat(worksheet.Cell("I6"), @"Рано");
            SetFormat(worksheet.Cell("C6"), @"приход");
            SetFormat(worksheet.Cell("D6"), @"уход");
            if (CheckWorkCount)
            {
                SetFormat(worksheet.Range("K5:K6"), @"Переработка");
            }
            return worksheet;
        }

        private IXLRow SetHolidayNoActionRow(IXLRow row)
        {
            SetFormat(row.Cell(3), "");
            SetFormat(row.Cell(4), "");
            SetFormat(row.Cell(5), "");
            SetFormat(row.Cell(6), "");
            SetFormat(row.Cell(7), "");
            SetFormat(row.Cell(8), "");
            SetFormat(row.Cell(9), "");
            return row;
        }

        private IXLWorksheet SetNotHolidayNoActionRow(IXLRow row, IXLWorksheet worksheet, string description)
        {
            SetFormat(worksheet.Range(row.Cell(3),row.Cell(4)),"Не зарегистрирован");
            SetFormat(row.Cell(5), description ?? "");
            SetFormat(row.Cell(6), TimeSpan.Zero);
            SetFormat(row.Cell(7), TimeSpan.Zero);
            SetFormat(row.Cell(8), 0);
            SetFormat(row.Cell(9), 0);
            return worksheet;
        }

        private static  TimeSpan RoundOverWork(TimeSpan overWork)
        {
            return overWork < TimeSpan.Zero 
                 ? TimeSpan.Zero
                 : TimeSpan.FromHours(((int) overWork.TotalHours))
                           .Add(TimeSpan.FromMinutes(overWork.Minutes < 30
                           ? TimeSpan.Zero.Minutes
                           : overWork.Minutes < 56
                                     ? 30
                                     : Hour.Minutes));
        }
        
        private XLWorkbook CreateWorksheetForUserPersonalReport(Person person, XLWorkbook workbook, DateTime startPeriod, DateTime endPeriod, List<int> holidays)
        {
            var worksheet = workbook.AddWorksheet(person.Name.Substring(0, person.Name
                                                                                                   .Length >= 31
                                                                                             ? 31
                                                                                             : person.Name
                                                                                                     .Length));
            worksheet = CreateHeaderWorkSheet(worksheet, startPeriod);
            /*init startup variable*/
            #region initialization
            var currentRow = 7; 
            var start = startPeriod; 
            var end = endPeriod.AddDays(1); 
            var payForLate = 0;
            var payForEarly = 0;
            TimeSpan earlyOutSum = TimeSpan.Zero,
                lateEnterSum = TimeSpan.Zero,
                overWorkSum = TimeSpan.Zero;
            #endregion

            while (start != end)
            {
                var row = worksheet.Row(currentRow);
                /*init holiday*/
                #region holiday
                bool holiday;
                if (start.DayOfWeek == DayOfWeek.Sunday || start.DayOfWeek == DayOfWeek.Saturday || holidays.Contains(start.Day))
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
                #endregion

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
                    if (!holiday) worksheet = SetNotHolidayNoActionRow(row, worksheet, description);
                    else row = SetHolidayNoActionRow(row);
                }
                else
                {
                    TimeSpan latecomer;
                    if (!TimeSpan.Equals(enter, TimeSpan.MinValue) && TimeSpan.Equals(outer, TimeSpan.MinValue))
                    {
                        SetFormat(row.Cell(3), $"{enter:hh}:{enter:mm}:{enter:ss}");
                        SetFormat(row.Cell(4), "");
                        SetFormat(row.Cell(5), description ?? "");
                        SetFormat(row.Cell(7), "Не отметился");
                        if (!holiday)
                        {
                            SetFormat(row.Cell(9), 150);
                            payForEarly += 150;
                            if (enter > person.Startday
                                              .Add(FifteenMinuts))
                            {
                                latecomer = enter - person.Startday;
                                lateEnterSum += latecomer;
                                SetFormat(row.Cell(6), $"{latecomer:hh}:{latecomer:mm}:{latecomer:ss}");
                                payForLate += 100;
                                if (enter < person.Startday
                                                  .Add(Hour))
                                {
                                    SetFormat(row.Cell(8), 100);
                                }
                                else
                                {
                                    SetFormat(row.Cell(8), "?");
                                    row.Cell(8)
                                       .Style
                                       .Fill
                                       .BackgroundColor = XLColor.Red;
                                }
                            }
                            else
                            {
                                SetFormat(row.Cell(6), TimeSpan.Zero);
                                SetFormat(row.Cell(8), 0);
                            }
                        }
                        else
                        {
                            SetFormat(row.Cell(9), 0);
                            SetFormat(row.Cell(6), TimeSpan.Zero);
                            SetFormat(row.Cell(8), 0);
                        }
                    }
                    else
                    {
                        TimeSpan earlyOut;
                        TimeSpan overWork;
                        if (TimeSpan.Equals(enter, TimeSpan.MinValue) && !TimeSpan.Equals(outer, TimeSpan.MinValue))
                        {
                            SetFormat(row.Cell(3), "");
                            SetFormat(row.Cell(4), $"{outer:hh}:{outer:mm}:{outer:ss}");
                            SetFormat(row.Cell(5), description ?? "");
                            SetFormat(row.Cell(6), "Не отметился");
                            if ((outer > person.Endday && CountWorker) || 
                                (outer > person.Endday.Add(Hour) && !CountWorker) && CheckWorkCount)
                            {
                                if (outer < person.Startday)
                                {
                                    overWork = (new TimeSpan(24, 0, 0).Subtract(person.Endday)).Add(outer);
                                    overWork = RoundOverWork(overWork);
                                    overWorkSum = overWorkSum.Add(overWork);
                                }
                                else
                                {
                                    overWork = outer - (person.Endday
                                                              .Add(Hour)
                                                              .Add(FifteenMinuts));
                                    overWork = RoundOverWork(overWork);
                                    overWorkSum = overWorkSum.Add(overWork);
                                }
                                SetFormat(row.Cell(11), $"{overWork:hh}:{overWork:mm}:{overWork:ss}");
                            }
                            else if (!holiday)
                            {
                                SetFormat(row.Cell(8), 100);
                                payForLate += 100;
                                if (outer < person.Endday && outer > person.Startday)
                                {
                                    earlyOut = person.Endday - -outer;
                                    earlyOutSum += earlyOut;
                                    SetFormat(row.Cell(7), $"{earlyOut:hh}:{earlyOut:mm}:{earlyOut:ss}");
                                    payForEarly += 100;
                                    if (outer > person.Endday
                                                      .Subtract(Hour))
                                    {
                                        SetFormat(row.Cell(9), 100);
                                    }
                                    else
                                    {
                                        SetFormat(row.Cell(9), "?");
                                        row.Cell(9).Style.Fill.BackgroundColor = XLColor.Red;
                                    }
                                }
                                else
                                {
                                    SetFormat(row.Cell(7), TimeSpan.Zero);
                                    SetFormat(row.Cell(9), 0);
                                }
                            }
                            else
                            {
                                SetFormat(row.Cell(8), 0);
                                SetFormat(row.Cell(7), TimeSpan.Zero);
                                SetFormat(row.Cell(9), 0);
                            }
                        }
                        else if (!TimeSpan.Equals(enter, TimeSpan.MinValue) && !TimeSpan.Equals(outer, TimeSpan.MinValue))
                        {
                            SetFormat(row.Cell(3), $"{enter:hh}:{enter:mm}:{enter:ss}");
                            SetFormat(row.Cell(4), $"{outer:hh}:{outer:mm}:{enter:ss}");
                            SetFormat(row.Cell(5), description ?? "");
                            if (holiday)
                            {
                                if (outer < enter)
                                {
                                    overWork = overWork = (new TimeSpan(24, 0, 0).Subtract(person.Endday)).Add(outer);
                                    overWork = RoundOverWork(overWork);
                                    overWorkSum = overWorkSum.Add(overWork);
                                }
                                else
                                {
                                    overWork = outer - enter;
                                    overWork = RoundOverWork(overWork);
                                    overWorkSum = overWorkSum.Add(overWork);
                                }
                                SetFormat(row.Cell(11), $"{overWork:hh}:{overWork:mm}:{overWork:ss}");
                            }
                            else
                            {
                                if (enter > person.Startday
                                                  .Add(FifteenMinuts))
                                {
                                    latecomer = enter - person.Startday
                                                              .Add(FifteenMinuts);
                                    lateEnterSum += latecomer;
                                    SetFormat(row.Cell(6), $"{latecomer:hh}:{latecomer:mm}:{latecomer:ss}");
                                    if (latecomer > Hour)
                                    {
                                        SetFormat(row.Cell(8), "?");
                                        row.Cell(8)
                                            .Style
                                            .Fill
                                            .BackgroundColor = XLColor.Red;
                                    }
                                    else
                                    {
                                        SetFormat(row.Cell(8), 100);
                                        payForLate += 100;
                                    }
                                }
                                else
                                {
                                    SetFormat(row.Cell(8), 0);
                                    SetFormat(row.Cell(6), TimeSpan.Zero);
                                }
                                if (outer < person.Endday && outer > person.Startday)
                                {
                                    earlyOut = person.Endday - outer;
                                    SetFormat(row.Cell(7), $"{earlyOut:hh}:{earlyOut:mm}:{earlyOut:ss}");
                                    earlyOutSum += earlyOut;
                                    if (earlyOut > Hour)
                                    {
                                        SetFormat(row.Cell(9), "?");
                                        row.Cell(9)
                                            .Style
                                            .Fill
                                            .BackgroundColor = XLColor.Red;
                                    }
                                    else
                                    {
                                        SetFormat(row.Cell(9), 150);
                                        payForEarly += 150;
                                    }
                                }
                                else
                                {
                                    SetFormat(row.Cell(7), TimeSpan.Zero);
                                    SetFormat(row.Cell(9), 0);
                                }
                                if (outer > person.Endday)
                                {
                                    overWork = outer - (person.Endday
                                                              .Add(Hour)
                                                              .Add(FifteenMinuts));
                                    overWork = RoundOverWork(overWork);
                                    overWorkSum = overWorkSum.Add(overWork);
                                    SetFormat(row.Cell(11), $"{overWork:hh}:{overWork:mm}:{overWork:ss}");
                                }
                            }
                        }
                    }
                }
                start = start.AddDays(1);
                currentRow++;
            }
            currentRow++;
            var summaryRow = worksheet.Row(currentRow);
            SetFormat(summaryRow.Cell(6), $"{lateEnterSum:hh}:{lateEnterSum:mm}:{lateEnterSum:ss}");
            SetFormat(summaryRow.Cell(7), earlyOutSum);
            SetFormat(summaryRow.Cell(8), payForLate);
            SetFormat(summaryRow.Cell(9), payForEarly);
            if (CheckWorkCount)
            {
                SetFormat(summaryRow.Cell(11), $"{overWorkSum:hh}:{overWorkSum:mm}:{overWorkSum:ss}");
            }
            var sumFullRange = worksheet.Range(currentRow + 1, 6, currentRow + 1, 7);
            var fullRange = worksheet.Range(currentRow + 1, 8, currentRow + 1, 9);
            sumFullRange.Merge();
            fullRange.Merge();
            var timeLow = earlyOutSum + lateEnterSum;
            var payLow = payForLate + payForEarly;
            SetFormat(sumFullRange, $"{timeLow:hh}:{timeLow:mm}:{timeLow:ss}");
            SetFormat(fullRange, payForLate + payForEarly);
            worksheet.Columns()
                     .AdjustToContents();
            worksheet.Rows()
                     .AdjustToContents();
            try
            {
                workbook.Save();
            }
            catch (Exception ex)
            {
                // ignored
            }
            return workbook;
        }

        private void AddPersonalReports(XLWorkbook workbook, List<Person> persons, DateTime startPeriod, DateTime endPeriod, List<int> holidays, ProgressBar progress)
        {
            var fifteenMinutes = new TimeSpan(0, 15, 0);
            var hour = new TimeSpan(1, 0, 0);
            progress.Value = 75;
            foreach (var person in persons)
            {
                workbook = CreateWorksheetForUserPersonalReport(person, workbook, startPeriod, endPeriod, holidays);
            }
        }

        public void GetReport(DateTime start, DateTime end, List<int> holidays, ProgressBar progressBar)
        {
            var workBook = GetReportForPeriod(start, end, out var persons, progressBar);
            AddPersonalReports(workBook, persons, start, end, holidays, progressBar);
            progressBar.Value = 100;
        }
    }
}
