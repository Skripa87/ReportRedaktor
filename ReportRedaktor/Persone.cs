using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportRedaktor
{
    public class Persone:IComparable
    {
        public string Name { get; set; }
        public List<Visit> VisitList { get; set; }
        public TimeSpan Startday { get; set; }
        public TimeSpan Endday { get; set; }
        public Persone(string name)
        {
            Name = name;
            VisitList = new List<Visit>();
        }

        public int CompareTo(object obj)
        {
            return string.Compare(Name, ((Persone)obj).Name);
        }
    }
}
