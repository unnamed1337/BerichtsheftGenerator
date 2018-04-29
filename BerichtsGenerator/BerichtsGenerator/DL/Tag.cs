using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BerichtsGenerator.DL
{
    public class Tag
    {
        public DateTime Datum { get; set; }
        public List<Buchung> Buchungen { get; set; }
        public Tag()
        {
            Buchungen = new List<Buchung>();
        }
        public Tag(DateTime date)
        {
            Datum = date;
            Buchungen = new List<Buchung>();
        }
    }
}
