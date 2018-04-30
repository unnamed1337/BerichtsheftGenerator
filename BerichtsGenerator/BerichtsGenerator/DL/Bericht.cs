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
            VerfasstAm = DateTime.Now;
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

                excelWorksheet.Row(9).Height = 32*2;
                excelWorksheet.Row(10).Height = 32*2;
                excelWorksheet.Row(11).Height = 32*2;
                excelWorksheet.Row(12).Height = 32*2;
                excelWorksheet.Row(13).Height = 32*2;

                excelWorksheet.Row(15).Height = 22 * 2;
                excelWorksheet.Row(16).Height = 22 * 2;



                excelWorksheet.Cells[1, 1].Value = "Name";
                excelWorksheet.Cells[1, 2].Value = Verfasser.Item1;
                excelWorksheet.Cells[2, 1].Value = "Vorname";
                excelWorksheet.Cells[2, 2].Value = Verfasser.Item2;
                excelWorksheet.Cells[4, 1].Value = "Ausbildungsberuf:";
                excelWorksheet.Cells[5, 1].Value = Beruf;
                
                excelWorksheet.Cells[1, 1, 5, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                //excelWorksheet.Cells[1, 4].Value = "Unternehmen";
                //excelWorksheet.Cells[3, 4].Value = Unternehmen;
                excelWorksheet.Cells[1, 3, 5, 5].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[1, 3, 5, 5].Merge = true;
                excelWorksheet.Cells[1, 3, 5, 5].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                excelWorksheet.Cells[1, 3, 5, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[1, 3, 5, 5].Value = "Unternehmen:\r\n\r\n" + Unternehmen;


                excelWorksheet.Cells[1, 6].Value = "Ausbildungsjahr ";
                excelWorksheet.Cells[1, 7].Value = To.ToString("yyyy");
                excelWorksheet.Cells[2, 6].Value = "Bericht Nr.:";
                excelWorksheet.Cells[2, 7].Value = BerichtNr.ToString();
                excelWorksheet.Cells[4, 6].Value = "Datum";
                excelWorksheet.Cells[4, 7].Value = VerfasstAm.ToString("dd.MM.yyyy");
                excelWorksheet.Cells[1, 6, 5, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin);


                excelWorksheet.Cells[8, 1].Value = "Von:\r\n"+ From.ToString("dd.MM.yyyy") + "\r\nBis:\r\n"+To.ToString("dd.MM.yyyy");
                excelWorksheet.Cells[8, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[8, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);

                excelWorksheet.Cells[8, 2, 8, 6].Merge = true;
                excelWorksheet.Cells[8, 2, 8, 6].Value = "Ausgeführte Arbeiten, Unterricht usw.";
                excelWorksheet.Cells[8, 2, 8, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[8, 2, 8, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[8, 2, 8, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                excelWorksheet.Cells[8, 7].Value = "Stunden";
                excelWorksheet.Cells[8, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[8, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[8, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Top;


                // Stunden Montag

                excelWorksheet.Cells[9, 1].Value = "Montag";
                excelWorksheet.Cells[9, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[9, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[9, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                excelWorksheet.Cells[9, 2, 9, 6].Merge = true;
                excelWorksheet.Cells[9, 2, 9, 6].Value = CreateBuchungsString(0);
                excelWorksheet.Cells[9, 2, 9, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[9, 2, 9, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                excelWorksheet.Cells[9, 2, 9, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                excelWorksheet.Cells[9, 7].Value = CreateStundenString(0);
                excelWorksheet.Cells[9, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[9, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[9, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Top;


                // Stunden Dienstag

                excelWorksheet.Cells[10, 1].Value = "Dienstag";
                excelWorksheet.Cells[10, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[10, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[10, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                excelWorksheet.Cells[10, 2, 10, 6].Merge = true;
                excelWorksheet.Cells[10, 2, 10, 6].Value = CreateBuchungsString(1);
                excelWorksheet.Cells[10, 2, 10, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[10, 2, 10, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                excelWorksheet.Cells[10, 2, 10, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                excelWorksheet.Cells[10, 7].Value = CreateStundenString(1);
                excelWorksheet.Cells[10, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[10, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[10, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Top;


                // Stunden Mittwoch

                excelWorksheet.Cells[11, 1].Value = "Mittwoch";
                excelWorksheet.Cells[11, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[11, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[11, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                excelWorksheet.Cells[11, 2, 11, 6].Merge = true;
                excelWorksheet.Cells[11, 2, 11, 6].Value = CreateBuchungsString(2);
                excelWorksheet.Cells[11, 2, 11, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[11, 2, 11, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                excelWorksheet.Cells[11, 2, 11, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                excelWorksheet.Cells[11, 7].Value = CreateStundenString(2);
                excelWorksheet.Cells[11, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[11, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[11, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Top;


                // Stunden Donnerstag

                excelWorksheet.Cells[12, 1].Value = "Donnerstag";
                excelWorksheet.Cells[12, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[12, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[12, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                excelWorksheet.Cells[12, 2, 12, 6].Merge = true;
                excelWorksheet.Cells[12, 2, 12, 6].Value = CreateBuchungsString(3);
                excelWorksheet.Cells[12, 2, 12, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[12, 2, 12, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                excelWorksheet.Cells[12, 2, 12, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                excelWorksheet.Cells[12, 7].Value = CreateStundenString(3);
                excelWorksheet.Cells[12, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[12, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[12, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Top;



                // Stunden Freitag

                excelWorksheet.Cells[13, 1].Value = "Freitag";
                excelWorksheet.Cells[13, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[13, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[13, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                excelWorksheet.Cells[13, 2, 13, 6].Merge = true;
                excelWorksheet.Cells[13, 2, 13, 6].Value = CreateBuchungsString(4);
                excelWorksheet.Cells[13, 2, 13, 6].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[13, 2, 13, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                excelWorksheet.Cells[13, 2, 13, 6].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                excelWorksheet.Cells[13, 7].Value = CreateStundenString(4);
                excelWorksheet.Cells[13, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[13, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                excelWorksheet.Cells[13, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Top;



                excelWorksheet.Cells[15, 1, 16, 4].Merge = true;
                excelWorksheet.Cells[15, 1, 16, 4].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[15, 1, 16, 4].Value = "Name und Unterschrift \r\ndes Auszubildenden";
                excelWorksheet.Cells[15, 1, 16, 4].Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                excelWorksheet.Cells[15, 5, 16, 7].Merge = true;
                excelWorksheet.Cells[15, 5, 16, 7].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                excelWorksheet.Cells[15, 5, 16, 7].Value = "gesehen:\r\n\r\nAusbilder";
                excelWorksheet.Cells[15, 5, 16, 7].Style.VerticalAlignment = ExcelVerticalAlignment.Top;



                FileInfo excelFile = new FileInfo(newFile);
                excel.SaveAs(excelFile);

                bool isExcelInstalled = Type.GetTypeFromProgID("Excel.Application") != null ? true : false;
                if (isExcelInstalled)
                {
                    System.Diagnostics.Process.Start(excelFile.ToString());
                }
            }

        }

        private string CreateBuchungsString(int i)
        {
            string output = "";
            foreach(Buchung buchung_ in Tagesbuchungen[i].Buchungen)
            {
                output += buchung_.Text+"\r\n";
            }
            return output;
        }

        private string CreateStundenString(int i)
        {
            string output = "";
            foreach (Buchung buchung_ in Tagesbuchungen[i].Buchungen)
            {
                output += buchung_.Stunden + "\r\n";
            }
            return output;
        }
    }
}
