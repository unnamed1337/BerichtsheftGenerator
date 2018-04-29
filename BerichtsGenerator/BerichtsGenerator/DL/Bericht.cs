using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BerichtsGenerator.DL
{
    public class Bericht
    {
        public int BerichtNr { get; set; }
        public Tuple<string,string> Verfasser { get; set; } //Vorname,Nachname
        public DateTime VerfasstAm { get; set; }
        public DateTime From { get; set; }
        public DateTime To { get; set; }
        public string Beruf { get; set; }
        public string Unternehmen { get; set; }
        public List<Tag> Tagesbuchungen { get; set; }

        public Bericht(Tuple<string,string>verfasser_,string beruf_,string unternehmen_,int berichtnr_,List<Tag>buchungen)
        {
            BerichtNr = berichtnr_;
            Verfasser = verfasser_;
            Unternehmen = unternehmen_;
            Beruf = beruf_;
            VerfasstAm = new DateTime();
            From = buchungen[0].Datum;
            To = buchungen[buchungen.Count - 1].Datum;
            Tagesbuchungen = buchungen;
        }

        public void ExportAsFile()
        {
            
        }
    }
}
