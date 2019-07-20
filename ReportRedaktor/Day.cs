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

        public Day(DateTime date)
        {
            Persones = new List<Persone>();
            Date = date;            
        }
    }
}
