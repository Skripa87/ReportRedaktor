using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportRedaktor
{
    public class Day
    {
        public DateTime Date { get; set; }
        public List<Persone> Persones { get; set; }

        public Day(string date)
        {
            Persones = new List<Persone>();
            Date = DateTime.TryParse(date, out var _date) == false
                 ? DateTime.Parse("01.01.1")
                 : _date;            
        }
    }
}
