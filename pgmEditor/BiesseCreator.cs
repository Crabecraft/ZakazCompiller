using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.IO;
using System.Windows.Forms;
using xNet;

namespace pgmEditor
{
    public static class BiesseCreator
    {
        private static string frezaName = "имя";
        private static string frezaSpeed = "вращение";
        private static string frezaPodacha = "подача";

        public static void Create(string patch, Form1 form)
        {
            loadFreza(form);
            Thread myBiesse = new Thread(new ParameterizedThreadStart(createBiesseProject));
            myBiesse.Start(new object[]{patch,form});
        }

        private static void createBiesseProject(object value)
        {
            string[] directories = Directory.GetDirectories(((object[])value)[0].ToString());
            string HomeBiessa = ((object[]) value)[0].ToString() + "Biesse";
            Form1 form = (Form1)((object[])value)[1];
            try
            {
                if(Directory.Exists(HomeBiessa))
                Directory.Delete(HomeBiessa, true);
            }
            catch
            {
                MessageBox.Show("Невозможно пересобрать!\nФайлы проекта используются в другом приложении.");
                return;
            }

            Directory.CreateDirectory(HomeBiessa);


            for (int i = 0; i < directories.Length; i++)
            {
                string[] stolDirectoris = Directory.GetDirectories(directories[i] + "\\");


                for (int j = 0; j < stolDirectoris.Length; j++)
                {
                    string[] tekDirectoryes = stolDirectoris[j].Split('\\');
                    string tekDirectory = tekDirectoryes[tekDirectoryes.Length-1];
                    string MaterialDirectory = HomeBiessa +"\\"+ tekDirectory;

                    if (!Directory.Exists(MaterialDirectory))
                    {
                        Directory.CreateDirectory(MaterialDirectory);
                        byte[] b = Properties.Resources.repair;
                        FileStream fs = new FileStream(MaterialDirectory + "\\repair.exe", FileMode.Create);
                        fs.Write(b, 0, b.Length);
                        fs.Close();

                        File.WriteAllLines(MaterialDirectory + "\\Apticles.csv", new string[] {""});
                    }

                    MaterialDirectory = MaterialDirectory + "\\";
                    string[] allFiles = Directory.GetFiles(stolDirectoris[j], "*.fkbpp");

                     

                    for (int x = 0; x < allFiles.Length; x++)
                    {
                        string[] temp = allFiles[x].Split('\\');
                        string[] tempName = temp[temp.Length - 1].Split('(');
                        

                        Detal деталь = new Detal();
                        деталь.Load(allFiles[x]);

                        string name = деталь.DX.ToString() + "x" + деталь.DY.ToString();

                        string BiesseData = деталь.BiesseData;
                        setFreza(ref BiesseData);
                        for (int z = 1 ; ;z++)
                        {
                            if (!File.Exists(MaterialDirectory + name +"("+ z.ToString() +")"+ ".cix"))
                            {
                                string nameFileProgramm = MaterialDirectory + name + "(" + z.ToString() + ")" + ".cix";

                                string[] alart = File.ReadAllLines(MaterialDirectory + "\\Apticles.csv");
                                ArrayList tempart = new ArrayList();
                                for (int g = 0; g < alart.Length; g++) tempart.Add(alart[g]);
                                

                                if (деталь.BiesseData.Contains("PARAM,NAME=TNM,VALUE=\"LAMA120\"") || !form.checkBox1.Checked)
                                    деталь.Article = деталь.Article.Replace("$rot", "0");
                                else
                                    деталь.Article = деталь.Article.Replace("$rot", "90");

                                tempart.Add(деталь.Article.Replace("$name", nameFileProgramm));


                                File.WriteAllLines(MaterialDirectory + "\\Apticles.csv", (string[])tempart.ToArray(typeof(string)));

                                    File.WriteAllText(nameFileProgramm, BiesseData);
                                if (деталь.ScmData != null)
                                    if (деталь.ScmData.Length > 0)
                                {
                                    string scmDirektory = MaterialDirectory + name + "(" + z.ToString() + ")";
                                    Directory.CreateDirectory(scmDirektory);
                                    scmDirektory += "\\";

                                    for (int c = 0; c < деталь.ScmData.Length; c++)
                                    {
                                        File.WriteAllText(scmDirektory + name + "(" + (c + 1).ToString() + ").xxl", деталь.ScmData[c]);
                                    }
                                }
                                break;
                            }
                        }
                    } 

               
                }

            }

            MessageBox.Show("Готово");
        }

        private static void loadFreza(Form1 form)
        {
            frezaName = form.button2.Text;

            Properties.Settings set = new Properties.Settings();

            string[] frezi = set.frezi.Split('$');

            for (int i = 0; i < frezi.Length; i++)
            {
                string[] opti = frezi[i].Split(';');
                if (opti[0] == frezaName)
                {
                    frezaSpeed = opti[2];
                    frezaPodacha = opti[1];
                    break;
                }
            }

        }

        private static void setFreza(ref string biesseData)
        {
           string[] newRout = biesseData.Split(new string[] { "\n\t" }, StringSplitOptions.RemoveEmptyEntries);

            for(int i=0; i < newRout.Length;i++)
                if (newRout[i] == "NAME=ROUTG")
                {
                    newRout[i + 15] = "PARAM,NAME=RSP,VALUE=" + frezaSpeed;
                    newRout[i + 17] = "PARAM,NAME=WSP,VALUE=" + frezaPodacha;
                    newRout[i + 59] = "PARAM,NAME=TNM,VALUE=\"" + frezaName + "\"";
                }

            biesseData = string.Join("\n\t", newRout);
        }
    }
}
