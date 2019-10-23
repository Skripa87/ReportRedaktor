using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reporter
{
    public class Visit
    {
        public DateTime Date { get; set; }
        public TimeSpan Enter { get; set; }
        public TimeSpan Outer { get; set; }

        public Visit(DateTime date, TimeSpan enter, TimeSpan outer)
        {
            Date = date;
            Enter = enter;
            Outer = outer;
        }
    }
}
