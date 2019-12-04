using System;

namespace Reporter
{
    public enum Direction { In, Out}
    public class WorkEvent
    {
        public DateTime Date { get; set; }
        public TimeSpan Time { get; set; }
        public string Point { get; set; }
        public Direction Direction { get; set; }
        public string Name { get; set; }
        public int Number { get; set; }
        public string UserName { get; set; }

        public WorkEvent(string date, string time, string point, string direction, string name, string number, string userName)
        {
            Date = DateTime.TryParse(date, out var ate)
                 ? ate
                 : DateTime.MinValue;
            Time = TimeSpan.TryParse(time, out var ime)
                 ? ime
                 : TimeSpan.MinValue;
            Point = point;
            Direction = direction.Contains("ы")
                      ? Direction.Out
                      : Direction.In;
            Name = name;
            Number = string.IsNullOrEmpty(number)
                   ? 0
                   : -1;
            UserName = userName;
        }
    }
}
