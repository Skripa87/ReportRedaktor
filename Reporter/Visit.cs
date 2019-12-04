using System;

namespace Reporter
{
    public class Visit
    {
        public DateTime Date { get; set; }
        public TimeSpan Enter { get; set; }
        public TimeSpan Outer { get; set; }
        public string Description { get; set; }
        
        public Visit(DateTime date, TimeSpan enter, TimeSpan outer, string description)
        {
            Date = date;
            Enter = enter;
            Outer = outer;
            Description = description;
        }
    }
}
