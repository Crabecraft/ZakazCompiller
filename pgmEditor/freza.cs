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
    public partial class freza : Form
    {
        Form1 form;
        public freza(Form1 form)
        {
            InitializeComponent();
            this.form = form;
        }

     

        private void freza_Load(object sender, EventArgs e)
        {
            Properties.Settings set = new Properties.Settings();
            string[] frezi = set.frezi.Split(new string[] {"$"},StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < frezi.Length; i++)
            {
                string[] opti = frezi[i].Split(';');
                dataGridView1.Rows.Add(new object[] { opti[0], opti[1], opti[2] });
            }
        }

     
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                int id = dataGridView1.CurrentRow.Index;
                if (!dataGridView1.Rows[id].Cells[0].Value.ToString().Trim().Equals(""))
                {
                    form.button2.Text = dataGridView1.Rows[id].Cells[0].Value.ToString().Trim();
                    form.WriteRotSetting();

                    string setSave = "";

                    for (int i = 0; i < dataGridView1.Rows.Count-1; i++)
                        if (!dataGridView1.Rows[i].Cells[0].Value.ToString().Trim().Equals(""))
                            setSave += dataGridView1.Rows[i].Cells[0].Value.ToString().Trim() + ";" + dataGridView1.Rows[i].Cells[1].Value.ToString().Trim() + ";" + dataGridView1.Rows[i].Cells[2].Value.ToString().Trim() + "$";

                    Properties.Settings set = new Properties.Settings();

                    set.frezi = setSave;
                    set.Save();

                    this.Close();
                }
            }
            catch { }
        }

        
    }
}
