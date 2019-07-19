using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ReportRedaktor
{
    public class Work_manager
    {
        public List<Work_event> GetWork_Events(string fileName)
        {
            var work_events = new List<Work_event>();
            string fileNameExcel = fileName;
            //XLWorkbook workbook = new XLWorkbook(fileNameExcel);
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileNameExcel, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts
                                                          .FirstOrDefault();
                SheetData sheetData = worksheetPart.Worksheet
                                                   .Elements<SheetData>()
                                                   .FirstOrDefault();
                var rows = sheetData.Elements<Row>().Cast<IXLRow>();
                foreach (Row row in sheetData.Elements<Row>())
                {
                    if (row.Elements<Cell>().Any(c=>c.InnerText != "") 
                      && int.TryParse(row.RowIndex,out var rowIndex) && rowIndex > 9) 
                    foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                    {
                        var x = cell.CellValue?.Text ?? "";
                    }
                }
            }
            //var workcheet = workbook.Worksheet(1);
        
            //int row = 9, column = 1;
            string datetime = "", point = "", direction = "", name = "", number = "", username = "";
            //while (true)
            //{
            //    if(string.IsNullOrEmpty(workcheet.Cell(row,column)
            //                                     .Value
            //                                     .ToString()))
            //    {
            //        break;
            //    }
            //    datetime = workcheet.Cell(row, column)
            //                        .Value
            //                        .ToString();
            //    column++;
            //    point = workcheet.Cell(row, column)
            //                     .Value
            //                     .ToString();
            //    column++;
            //    direction = workcheet.Cell(row, column)
            //                         .Value
            //                         .ToString();
            //    column++;
            //    name = workcheet.Cell(row, column)
            //                    .Value
            //                    .ToString();
            //    column++;
            //    number = workcheet.Cell(row, column)
            //                      .Value
            //                      .ToString();
            //    if (string.IsNullOrEmpty(number))
            //    {
            //        column++;
            //        username = workcheet.Cell(row, column)
            //                            .Value
            //                            .ToString();
            //    }
            //    else
            //    {
            //        username = "";
            //    }
            //    var work_event = new Work_event(datetime,point,direction,name,number,username);
            //    column = 1;
            //    row++;
            //    work_events.Add(work_event);
            //}
            return work_events;
        }
    }
}
