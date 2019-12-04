using System;
using System.Collections.Generic;
using System.Linq;

namespace Reporter
{
    public class Person:IComparable
    {
        public string Name { get; set; }
        public List<Visit> VisitList { get; set; }
        public TimeSpan Startday { get; set; }
        public TimeSpan Endday { get; set; }
        public Person(string name)
        {
            Name = name;
            VisitList = new List<Visit>();
        }

        public int CompareTo(object obj)
        {
            return string.CompareOrdinal(Name, ((Person)obj).Name);
        }

        public void SetWorkTime(List<string> fiveClockEndDayWorkers)
        {
            if (fiveClockEndDayWorkers.Any(f => Name.ToLower()
                                                    .Contains(f)))
            {
                Startday = TimeSpan.Parse("08:00:00");
                Endday = TimeSpan.Parse("17:00:00");
            }
            else
            {
                Startday = TimeSpan.Parse("09:00:00");
                Endday = TimeSpan.Parse("18:00:00");
            }
        }
    }
}
