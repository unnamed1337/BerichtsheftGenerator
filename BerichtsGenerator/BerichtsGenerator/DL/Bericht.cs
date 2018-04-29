using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
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
            string templatePath = "../../../../Bericht.xls";
            string newFile = "Bericht_"+BerichtNr+".xls";

            File.Copy(templatePath, newFile);
            // copy Templae

            //var fileinfo = new FileInfo(newFile);
            //fileinfo = new FileInfo(templatePath);
            //if (fileinfo.Exists)
            //{
            //    using (ExcelPackage p = new ExcelPackage(fileinfo))
            //    {
            //        ExcelWorksheet ws = p.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Sheet1");
            //        ws.Cells[9, 2].Value = Tagesbuchungen[0].Buchungen[0];
            //        //ws.Cells[1, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            //        //ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(184, 204, 228));
            //        //ws.Cells[1, 1].Style.Font.Bold = true;
            //        p.Save();
            //    }

            //}
        }
    }
}
