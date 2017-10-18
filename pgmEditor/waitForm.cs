using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace pgmEditor
{
    public partial class waitForm : Form
    {
        public int progress;
        public waitForm()
        {
            InitializeComponent();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (progress >= 100)
            {
                this.Close();
                return;
            }

            progressBar1.Value = progress;
        }
    }
}
