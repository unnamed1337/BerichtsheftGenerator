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
            //string templatePath = "../../../../Bericht.xls";
            string newFile = "Bericht_"+BerichtNr+".xlsx";

            var file = new FileInfo(newFile);
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("Bericht "+BerichtNr);
               
                var excelWorksheet = excel.Workbook.Worksheets["Bericht " + BerichtNr];

                excelWorksheet.Column(1).Width = 25.25/2;
                excelWorksheet.Column(2).Width = 18/2;
                excelWorksheet.Column(3).Width = 14.96/2;
                excelWorksheet.Column(4).Width = 31.34/2;
                excelWorksheet.Column(5).Width = 26.07/2;
                excelWorksheet.Column(6).Width = 31.91/2;
                excelWorksheet.Column(7).Width = 21.64/2;



                excelWorksheet.Cells[1, 1].Value = "Name";
                excelWorksheet.Cells[1, 2].Value = Verfasser.Item1;
                excelWorksheet.Cells[2, 1].Value = "Vorname";
                excelWorksheet.Cells[2, 2].Value = Verfasser.Item2;
                excelWorksheet.Cells[4, 1].Value = "Ausbildungsberuf:";
                excelWorksheet.Cells[5, 1].Value = Beruf;
                
                excelWorksheet.Cells[1, 1, 5, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                excelWorksheet.Cells[1, 4].Value = "Unternehmen";
                excelWorksheet.Cells[3, 4].Value = Unternehmen;
                excelWorksheet.Cells[1, 3, 5, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                excelWorksheet.Cells[1, 6].Value = "Ausbildungsjahr ";
                excelWorksheet.Cells[2, 6].Value = "Bericht Nr.:";
                excelWorksheet.Cells[2, 7].Value = BerichtNr.ToString();
                excelWorksheet.Cells[4, 6].Value = "Datum";
                excelWorksheet.Cells[1, 6, 5, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin);




                FileInfo excelFile = new FileInfo(newFile);
                excel.SaveAs(excelFile);

                bool isExcelInstalled = Type.GetTypeFromProgID("Excel.Application") != null ? true : false;
                if (isExcelInstalled)
                {
                    System.Diagnostics.Process.Start(excelFile.ToString());
                }
            }

        }
    }
}
