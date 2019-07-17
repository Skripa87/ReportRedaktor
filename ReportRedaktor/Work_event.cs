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
        public DateTime DateAndTime { get; set; }
        public string Point { get; set; }
        public Direction Direction { get; set; }
        public string Name { get; set; }
        public int Number { get; set; }
        public string UserName { get; set; }

        public Work_event(string dateAndTime, string point, string direction, string name, string number, string userName)
        {
            var date = new DateTime();
            DateAndTime = DateTime.TryParse(dateAndTime, out date)
                        ? date
                        : DateTime.MinValue;
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
