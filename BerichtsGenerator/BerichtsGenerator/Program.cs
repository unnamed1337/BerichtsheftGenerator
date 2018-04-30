using BerichtsGenerator.DL;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BerichtsGenerator
{
    static class Program
    {
        /// <summary>
        /// Der Haupteinstiegspunkt für die Anwendung.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
        public static List<Bericht> ImportBuchungen(string FilePath,Tuple<string,string>Verfasser,string unternehmen,string beruf,int berichtnr)
        {
            List<Bericht> Berichte = new List<Bericht>();
            List<Tag> AlleTage = new List<Tag>(); 

            StreamReader reader = new StreamReader(File.OpenRead(FilePath));
            string headerLine = reader.ReadLine();
            while (!reader.EndOfStream)
            {
                string line = reader.ReadLine();
                if (!String.IsNullOrWhiteSpace(line))
                {
                    string[] values = line.Split(';');
                    Tag tag_ = new Tag();
                    for (int i = 0; i < values.Length; i++)
                    {
                        if(i == 0)
                        {
                            tag_ = new Tag(Convert.ToDateTime(values[i]));
                        }
                        else if(!string.IsNullOrEmpty(values[i]))
                        {
                            string buchungString = values[i];
                            Tuple<string, double> buchung = SplitBuchung(buchungString); 
                            Buchung buchung_ = new Buchung(buchung.Item1, buchung.Item2);
                            tag_.Buchungen.Add(buchung_);
                        }
                    }
                    AlleTage.Add(tag_);
                }
            }
            int count = berichtnr;
            for(int i = 0; i <= AlleTage.Count - 5; i = i + 4)
            {
                List<Tag> tageTemp = new List<Tag>();
                tageTemp.Add(AlleTage[i]);
                tageTemp.Add(AlleTage[i + 1]);
                tageTemp.Add(AlleTage[i + 2]);
                tageTemp.Add(AlleTage[i + 3]);
                tageTemp.Add(AlleTage[i + 4]);
                Bericht berichtTemp = new Bericht(Verfasser, beruf, unternehmen, count, tageTemp);
                Berichte.Add(berichtTemp);
                count++;
            }



            return Berichte;
        }

        private static Tuple<string,double> SplitBuchung(string BuchungsString)
        {
            string aufgabe = BuchungsString;
            string stundenString = "";
            double stunden = 1;

            for(int i = BuchungsString.Length - 1; i >= 0; i--)
            {
                if (BuchungsString[i] == '-')
                {
                    break;
                }
                stundenString = BuchungsString[i] + stundenString;
            }

            aufgabe = aufgabe.Replace(" -" + stundenString, "");
            stundenString = stundenString.Replace("Stunden", "");
            stundenString = stundenString.Replace(" ", "");
            
            try
            {
                stunden = Convert.ToDouble(stundenString);
            }
            catch
            {
                MessageBox.Show(BuchungsString);
            }
            return new Tuple<string, double>(aufgabe, stunden);
        }
    }
}
