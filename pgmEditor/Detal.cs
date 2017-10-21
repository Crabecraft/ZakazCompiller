using System;
using xNet;
using System.Xml.Serialization;
using System.IO;
using System.Windows.Forms;

namespace pgmEditor
{
    public class Detal
    {
        public float DX, DY, DZ;

        public string BiesseData = "";
        public string Article;
        public string StolParam;
        public string DetalParam;
        public string[] ScmData;
        public string[] BiesseListData;

        public string кромка = "0,0,0,0";

        public void setSize(string DX, string DY, string DZ)
            {
                this.DX = float.Parse(DX);
                this.DY = float.Parse(DY);
                this.DZ = float.Parse(DZ);
            }

        public string generate()
            {
               

                string temp = File.ReadAllText(Application.StartupPath +  "\\BiessePrograms\\Macros\\Шапка.bmac") + "\n"; ;

                if (BiesseListData != null)
                    for (int i = 0; i < BiesseListData.Length; i++)
                    {
                        if (BiesseListData[i] != "")
                        {
                            string param = "";
                            string nameProg = "";
                            if (BiesseListData[i].Contains("$"))
                            {
                                param = "$" + BiesseListData[i].Split('$')[1];
                                nameProg = BiesseListData[i].Split('$')[0];
                            }
                            else
                                nameProg = BiesseListData[i];

                            string macrosPatch = Application.StartupPath + "\\BiessePrograms\\Macros\\" + nameProg + ".bmac";
                            if (!File.Exists(macrosPatch)) { MessageBox.Show("Файл: " + macrosPatch + "\n не найден."); return null; }



                            string programm = setMINRoutg(macrosPatch);



                            if (param != "")
                            {
                                string[] values = param.Split('=')[1].Split(new string[] { "_" }, StringSplitOptions.RemoveEmptyEntries);
                                for (int j = 0; j < values.Length; j++)
                                {
                                    string tekMakro = programm;
                                    temp += tekMakro.Replace(param.Split('=')[0], values[j]) + "\n";
                                }

                            }
                            else {
                                temp += programm + "\n"; }
                        }
                    }

            string[] sep = new string[] { ";" };
            string[] paramMass = new string[1];

            if (StolParam != null)
            {
                paramMass = StolParam.Split(sep, StringSplitOptions.RemoveEmptyEntries);

                if (paramMass != null)
                    for (int i = 0; i < paramMass.Length; i++)
                        temp = temp.Replace(paramMass[i].Split('=')[0].Trim(), paramMass[i].Split('=')[1].Trim());
            }

            if (DetalParam != null)
            {
                paramMass = DetalParam.Split(sep, StringSplitOptions.RemoveEmptyEntries);
                if (paramMass != null)
                    for (int i = 0; i < paramMass.Length; i++)
                        temp = temp.Replace(paramMass[i].Split('=')[0].Trim(), paramMass[i].Split('=')[1].Trim());
            }
                temp = temp.Replace("$lpx", DX.ToString()).Replace("$lpy", DY.ToString()).Replace("$lpz", DZ.ToString());

                return temp;
            }

        public void Save(string patch)
            {
                try
                {
                    File.Delete(patch);
                    XmlSerializer formatter = new XmlSerializer(typeof(Detal));
                    using (FileStream fs = new FileStream(patch, FileMode.OpenOrCreate))
                    {
                        formatter.Serialize(fs, this);
                    }
                }
                catch { MessageBox.Show("Не получилось сохранить, возможно выбранный файл используется другой программой."); }
            }

        public string setMINRoutg(string patch)
        {
            if (DX * DY < 80001f)
            {

                string[] temp = File.ReadAllText(patch).Split(new string[] { "BEGIN MACRO" }, StringSplitOptions.RemoveEmptyEntries);

                for (int i = 0; i < temp.Length; i++)
                    if (temp[i].Contains("NAME=ROUTG"))
                    {
                        if (temp[i].Contains("PARAM,NAME=GID,VALUE=\"G1001.1001\""))
                        {
                            if (temp[i].Contains("PARAM,NAME=THR,VALUE=1"))
                            {
                                string repStr = temp[i].Split(new string[] { "PARAM,NAME=VTR,VALUE=" }, StringSplitOptions.RemoveEmptyEntries)[1];
                                repStr = repStr.Split(new string[] { "PARAM,NAME=INCSTP,VALUE=" }, StringSplitOptions.RemoveEmptyEntries)[0];
                                string[] values = repStr.Split(new string[] { "PARAM,NAME=DVR,VALUE=" }, StringSplitOptions.RemoveEmptyEntries);

                                int value = int.Parse(values[0].Trim());
                                value++;

                                string newLastProhod = "0.7";
                                string newColProhodon = value.ToString();

                                string returnStr = File.ReadAllText(patch);

                                temp[i] = temp[i].Replace("PARAM,NAME=VTR,VALUE=" + values[0], " PARAM,NAME=VTR,VALUE=" + newColProhodon + "\n");
                                temp[i] = temp[i].Replace("PARAM,NAME=DVR,VALUE=" + values[1], " PARAM,NAME=DVR,VALUE=" + newLastProhod + "\n");
                                return string.Join("BEGIN MACRO\n", temp);
                            }
                        }
                    }
            }
            return File.ReadAllText(patch);
        }

        public void Load(string patch)
            {
                this.Article = "";
                this.DetalParam = "";
                this.StolParam = "";
                try
                {
                    XmlSerializer formatter = new XmlSerializer(typeof(Detal));
                    using (FileStream fs = new FileStream(patch, FileMode.OpenOrCreate))
                    {
                        Detal temp = (Detal)formatter.Deserialize(fs);
                        this.DX = temp.DX;
                        this.DY = temp.DY;
                        this.DZ = temp.DZ;
                        this.BiesseData = temp.BiesseData;
                        this.кромка = temp.кромка;
                        this.Article = temp.Article;
                        this.DetalParam = temp.DetalParam;
                        this.StolParam = temp.StolParam;
                        this.ScmData = new string[temp.ScmData.Length];
                        this.BiesseListData = new string[temp.BiesseListData.Length];
                        for (int i = 0; i < temp.ScmData.Length; i++)
                            this.ScmData[i] = temp.ScmData[i];
                        for (int i = 0; i < temp.BiesseListData.Length; i++)
                            this.BiesseListData[i] = temp.BiesseListData[i];
                    }
                }
                catch { }
            }        

        public void set(string ParamDetal,string ParamStol)
        {
            Article = Article.Replace("$lpx", DX.ToString()).Replace("$lpy", DY.ToString()).Replace("$lpz", DZ.ToString());            
            Article = Article.Replace("_art", "_" + Guid.NewGuid().ToString());

            //if (ScmData != null)
            //    for (int i = 0; i < ScmData.Length; i++)
            //    {
            //        ScmData[i] = ScmData[i].Replace("$lpx", DX.ToString()).Replace("$lpy", DY.ToString()).Replace("$lpz", DZ.ToString());
                    
            //        for (int j = 0; j < massParam.Length; j++)
            //       {
            //        string[] value = massParam[j].Split('=');
            //        ScmData[i] = ScmData[i].Replace(value[0].Replace(" ", ""), value[1].Replace(" ", ""));
            //       }
            //    }
        }

    }
}
