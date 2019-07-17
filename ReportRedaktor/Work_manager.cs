using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ReportRedaktor
{
    public class Work_manager
    {
        public List<Work_event> GetWork_Events(string fileName)
        {
            var work_events = new List<Work_event>();
            string fileNameExcel = fileName;
            
            var workcheet = workbook.Worksheet(1);
            int row = 9, column = 1;
            string datetime = "", point = "", direction = "", name = "", number = "", username = "";
            while (true)
            {
                if(string.IsNullOrEmpty(workcheet.Cell(row,column)
                                                 .Value
                                                 .ToString()))
                {
                    break;
                }
                datetime = workcheet.Cell(row, column)
                                    .Value
                                    .ToString();
                column++;
                point = workcheet.Cell(row, column)
                                 .Value
                                 .ToString();
                column++;
                direction = workcheet.Cell(row, column)
                                     .Value
                                     .ToString();
                column++;
                name = workcheet.Cell(row, column)
                                .Value
                                .ToString();
                column++;
                number = workcheet.Cell(row, column)
                                  .Value
                                  .ToString();
                if (string.IsNullOrEmpty(number))
                {
                    column++;
                    username = workcheet.Cell(row, column)
                                        .Value
                                        .ToString();
                }
                else
                {
                    username = "";
                }
                var work_event = new Work_event(datetime,point,direction,name,number,username);
                column = 1;
                row++;
                work_events.Add(work_event);
            }
            return work_events;
        }
    }
}
