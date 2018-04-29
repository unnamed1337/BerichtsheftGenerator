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

        }
    }
}
