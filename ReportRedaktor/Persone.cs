using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportRedaktor
{
    public class Persone
    {
        public string Name { get; set; }
        public List<Visit> VisitList { get; set; }
        public Persone(string name)
        {
            Name = name;
            VisitList = new List<Visit>();
        }
    }
}
