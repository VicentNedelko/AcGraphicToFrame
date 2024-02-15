using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AcGraphicToFrame.Forms
{
    public partial class Info : Form
    {
        public Info()
        {
            InitializeComponent();
        }

        public void SetTextInfo(string text)
        {
            label1.Text = text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
