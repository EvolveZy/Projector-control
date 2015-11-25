using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CloverControl
{
    public partial class yz : Form
    {
        public yz()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(kl.Text.Trim()=="hello clover")
            {
                this.Hide();
                Form Setup = new Setup();
                Setup.ShowDialog();
            }
            else
            {
                MessageBox.Show("亲，口令不对哦");
            }
        }

        private void yz_Activated(object sender, EventArgs e)
        {
            kl.Focus();
        }
    }
}
