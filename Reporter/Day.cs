using System;
using System.Collections.Generic;

namespace Reporter
{
    public class Day
    {
        public DateTime Date { get; set; }
        public List<Person> Persons { get; set; }

        public Day(DateTime date)
        {
            Persons = new List<Person>();
            Date = date;            
        }
    }
}
