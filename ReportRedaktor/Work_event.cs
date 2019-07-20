using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportRedaktor
{
    public enum Direction { IN, OUT}
    public class Work_event
    {
        public DateTime Date { get; set; }
        public TimeSpan Time { get; set; }
        public string Point { get; set; }
        public Direction Direction { get; set; }
        public string Name { get; set; }
        public int Number { get; set; }
        public string UserName { get; set; }

        public Work_event(string date, string time, string point, string direction, string name, string number, string userName)
        {
            Date = DateTime.TryParse(date, out var _date)
                        ? _date
                        : DateTime.MinValue;
            Time = TimeSpan.TryParse(time, out var _time)
                        ? _time
                        : TimeSpan.MinValue;
            Point = point;
            Direction = direction.Contains("ы")
                      ? Direction.OUT
                      : Direction.IN;
            Name = name;
            Number = string.IsNullOrEmpty(number)
                   ? 0
                   : -1;
            UserName = userName;
        }
    }
}
