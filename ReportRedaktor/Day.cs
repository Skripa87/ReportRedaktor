using System;
using System.Collections.Generic;

namespace Reporter
{
    public class Day
    {
        public DateTime Date { get; set; }
        public List<Persone> Persons { get; set; }

        public Day(DateTime date)
        {
            Persons = new List<Persone>();
            Date = date;            
        }
    }
}
