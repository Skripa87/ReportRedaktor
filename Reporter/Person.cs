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

        public void SetWorkTime(List<string> whoStartNine)
        {
            if (whoStartNine.Any(f => Name.ToLower()
                                          .Contains(f)))
            {
                Startday = new TimeSpan(9, 0, 0);
                Endday = new TimeSpan(18, 0, 0);
            }
            else
            {
                Startday = new TimeSpan(8, 0, 0);
                Endday = new TimeSpan(17, 0, 0);
            }
        }
    }
}
