using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportRedaktor
{
    public class PersonalReport
    {
        public string Name { get; set; }
        public List<Visit> Visits { get; set; }

        public PersonalReport()
        {
            Visits = new List<Visit>();
        }
    }
}
