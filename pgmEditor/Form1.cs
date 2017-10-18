using System;
using System.Windows.Forms;
using System.Collections;
using System.Drawing;
using System.Threading;
using System.IO;
using System.Diagnostics;
using System.Linq;
using xNet;

namespace pgmEditor
{
    public partial class Form1 : Form
    {
        string папка = "";
        string test;
        public string папкаПроекта = "";
        public bool result;
       
        
        public Form1()
        {
            InitializeComponent();
            button1.AllowDrop = this.AllowDrop = true;
            load_form();
        }

        private void listBox1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && e.Effect == DragDropEffects.Move)
            {
                string name = ((string[])e.Data.GetData("FileNameW"))[0];

                string[] mass_name = name.Split('.');
                string TEST = mass_name[mass_name.Length - 1];
                if ((TEST == "xls") || (TEST == "xlsm"))
                    {
                      //if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                      //{
                        //папка = folderBrowserDialog1.SelectedPath + "\\";
                        папка = "C:\\BNest_Projects\\";
                        waitForm wait = new waitForm();
                        excelReader.CreateProject(name, папка, wait,this);
                        wait.ShowDialog();
                      //}
                    }
            }
        }

        private void listBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) &&
                ((e.AllowedEffect & DragDropEffects.Move) == DragDropEffects.Move))
                e.Effect = DragDropEffects.Move;
        }

        public void load_form()
        {
            string homePatch = "C:\\BNest_Projects";
            listBox1.Items.Clear();
            ReadRotSetting();

            if (!Directory.Exists(homePatch))
            {
                Directory.CreateDirectory(homePatch); return;
            }

            

            var dirs = Directory.GetDirectories(homePatch);
            if (dirs == null) return;


            var query = from dir in dirs
                        let tmp = new FileInfo(dir)
                        orderby tmp.CreationTime
                        select new { folder = dir, Date = tmp.CreationTime };
            ArrayList projects = new ArrayList();

            foreach (var a in query)
            {
                string[] name = a.folder.Split(new string[] { "\\" }, StringSplitOptions.RemoveEmptyEntries);
                //listBox1.Items.Add(name[name.Length - 1]);
                projects.Add(name[name.Length - 1]);
            }


            for (int i = projects.Count - 1; i >= 0; i--)
            {
                string[] name = projects[i].ToString().Split(new string[] { "\\" }, StringSplitOptions.RemoveEmptyEntries);
                //string[] cvet =  File.ReadAllLines(projects[i] + "\\cvet.txt");
                listBox1.Items.Add(name[name.Length - 1]);
            }

                     //for (int i = 0; i < projects.Count; i++)
                     //{
                     //    listBox1.Items.Add(projects[i].ToString());
                     //}
                
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //if (папкаПроекта == "")
            //{
            //    if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            //    {
            //        setLabel(folderBrowserDialog1.SelectedPath + "\\");
            //    }
            //}
            //if (папкаПроекта == "") return;

            if (listBox1.Items.Count < 1) return;

            if (listBox1.SelectedIndex > -1)
                BiesseCreator.Create("C:\\BNest_Projects\\" + listBox1.Items[listBox1.SelectedIndex] + "\\",this);
        }

        public void setLabel(string value)
        {
            папкаПроекта = value;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (label1.Text != папкаПроекта)
            {
                label1.Text = папкаПроекта;
                load_form();
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            WriteRotSetting();
        }

        public void WriteRotSetting()
        {
            File.Delete("C:\\ZacazCompillerSettings.txt");
            File.AppendAllLines("C:\\ZacazCompillerSettings.txt", new string[]{ checkBox1.Checked.ToString(), button2.Text });
        }

        

        public void ReadRotSetting()
        {
            if (!File.Exists("C:\\ZacazCompillerSettings.txt")) return;
            string[] setting = File.ReadAllLines("C:\\ZacazCompillerSettings.txt");

            checkBox1.Checked = bool.Parse(setting[0]);
            button2.Text = setting[1];
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            if (result) { load_form(); result = false; }
        }


        private void listBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                listBox1.ClearSelected();
                button1.Enabled = false;
            }
        }

        private void listBox1_MouseClick(object sender, MouseEventArgs e)
        {
            if (listBox1.Text != "" && button1.Enabled == false) button1.Enabled = true;
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer", "C:\\BNest_Projects\\" + listBox1.Items[listBox1.SelectedIndex]);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            freza frez = new freza(this);
            frez.ShowDialog();
        }

   
  
      

 
       
    }
}
