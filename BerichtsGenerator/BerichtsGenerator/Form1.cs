using BerichtsGenerator.DL;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BerichtsGenerator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label4_Click(object sender, EventArgs e)
        {
            label5_Click(null, null);
        }

        private void label5_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            label5.Text = openFileDialog1.FileName;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            bool failed = false;
            try
            {
                List<Bericht> Berichte  = Program.ImportBuchungen(openFileDialog1.FileName,new Tuple<string, string>(textBox2.Text, textBox3.Text),textBox1.Text,textBox4.Text,Convert.ToInt32(numericUpDown1.Value));
                foreach(Bericht tmpBericht in Berichte)
                {
                    tmpBericht.ExportAsFile();
                }
            }
            catch (Exception ex)
            {
                failed = true;
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (!failed)
                {
                    MessageBox.Show("Berichte wurden erstellt", "Done");
                }
            }
            //SaveSettings
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //tryLoadSettings
        }
    }
}
