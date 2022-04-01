using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Catering
{
    public partial class Signatures : Form
    {
        public String stringBox8;
        public String stringBox9;
        public String stringBox10;
        public String stringDate2;
        public String stringBox30;
        public String stringBox31;
        public String stringBox32;

        public Signatures()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            stringBox8 = textBox8.Text;
            stringBox9 = textBox9.Text;
            stringBox10 = textBox10.Text;
            stringBox30 = textBox30.Text;
            stringBox31 = textBox31.Text;
            stringBox32 = textBox32.Text;
            stringDate2 = dateTimePicker2.Text;

            this.Close();
        }
    }
}
