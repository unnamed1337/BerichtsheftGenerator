using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BerichtsGenerator.DL
{
    public class Buchung
    {
        public string Text { get; set; }
        public double Stunden { get; set; }
        public Buchung(string text_, double stunden_)
        {
            Text = text_;
            Stunden = stunden_;
        } 
    }
}
