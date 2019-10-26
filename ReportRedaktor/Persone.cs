using System;
using System.Collections.Generic;

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
    }
}
