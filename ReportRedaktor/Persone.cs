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
        public List<string> EnteredList { get; set; }
        public List<string> GetOutList { get; set; }
        public Persone(string name)
        {
            Name = name;
            EnteredList = new List<string>();
            GetOutList = new List<string>();
        }
    }
}
