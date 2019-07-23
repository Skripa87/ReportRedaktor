using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Security.Cryptography;
using OfficeOpenXml;
using ClosedXML.Excel;

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
                for(int row = 9; row < end.Row; row++)
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

        private List<Persone> GetPersones()
        {
            var persones = new List<Persone>();
            var work_events = GetWork_Events();
            while(work_events.Count > 0)
            {
                var persone = new Persone(work_events.FirstOrDefault().UserName);
                var personeEvents = work_events.FindAll(e => string.Equals(e.UserName, persone.Name));
                work_events.RemoveAll(e => string.Equals(e.UserName, persone.Name));
                while (personeEvents.Count > 0)
                {
                    var curent_date = personeEvents.FirstOrDefault().Date;
                    var events = personeEvents.FindAll(p => DateTime.Equals(p.Date, curent_date));
                    personeEvents.RemoveAll(p => DateTime.Equals(p.Date, curent_date));
                    var enter = events.Count(e => e.Direction == Direction.IN) > 0
                              ? events.FindAll(e => e.Direction == Direction.IN)
                                      .Min(e => e.Time)
                              : TimeSpan.MinValue;
                    var outer = events.Count(e => e.Direction == Direction.OUT) > 0
                              ? events.FindAll(e => e.Direction == Direction.OUT)
                                      .Max(m => m.Time)
                              :TimeSpan.MinValue;
                    var visit = new Visit(curent_date, enter, outer);
                    persone.VisitList.Add(visit);
                }
                persones.Add(persone);
            }
            return persones;
        }

        public XLWorkbook GetReportForPeriod(DateTime start, DateTime end)
        {
            var persones = GetPersones();
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
            worksheet.Cell("B3")
                     .Value = "Период";
            worksheet.Cell("C3")
                     .Value = start.Month
                                   .ToString();
            range = worksheet.Range("A5:A6");
            range.Merge()
                 .SetValue<string>("Дата");
            range = worksheet.Range("B5:B6");
            range.Merge()
                 .SetValue<string>("ФИО");
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
            workbook.SaveAs(newFileName);
            return workbook;
        }

        public 

    }
}
