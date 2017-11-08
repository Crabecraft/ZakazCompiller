using System;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.IO;

namespace pgmEditor
{
    public static class excelReader
    {
      static Thread loadThread;
      static string[] RusEng;

      static void Alfavit()
      {
          RusEng = new string[]
          {  
              "а=a","б=b","в=v", "г=g", "д=d", "е=e","ё=e", "ж=j","з=z","и=i","й=y'","к=k","л=l","м=m","н=n",
              "о=o","п=p","р=r", "с=s", "т=t", "у=u","ф=f", "х=h", "ц=c", "ч=ch", "ш=sh", "щ=sch", "ъ= ",
              "ы=i'",  "ь='",  "э=e", "ю=yu",  "я=ya", "А=A","Б=B", "В=V", "Г=G","Д=D", "Е=E","Ё=E", "Ж=J",
              "З=Z","И=I","Й=Y'",  "К=K", "Л=L",  "М=M","Н=N", "О=O","П=P","Р=R", "С=S","Т=T", "У=U", "Ф=F",
              "Х=H", "Ц=C",  "Ч=Ch",   "Ш=Sh","Щ=Sch", "Ъ= ", "Ы=I'","Ь='", "Э=E", "Ю=Yu", "Я=Ya"
          };
      }

      public static bool convertShablon(string patch, stol стол, Form1 form)
     {
          detal[] детали = (detal[])стол.детали.ToArray(typeof (detal));

         for (int i = 0; i < детали.Length; i++)
         {
             string homeCvet = patch + repStroka(детали[i].cvet);
             if(!Directory.Exists(homeCvet))
                 Directory.CreateDirectory(homeCvet);
             //homeCvet+="\\";

             Detal деталь = new Detal();
             деталь.Article = "1;_art;legno;$name;$lpx;$lpy;$lpz;$rot;0;0;0";




             if (детали[i].Programma == "" || детали[i].Programma == null)
             {
                 деталь.BiesseListData = new string[] { "Контур" };
                 //if (getPL(детали[i])) деталь.BiesseListData = new string[] { "КонтурMIN", "Контур" };
             }
             else
             {
                 //if (getPL(детали[i])) детали[i].Programma = "КонтурMIN;" + детали[i].Programma;
                 деталь.BiesseListData = детали[i].Programma.Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
             }

             деталь.setSize(детали[i].DX, детали[i].DY, детали[i].DZ);
             деталь.StolParam = детали[i].ParamStol;
             деталь.DetalParam = детали[i].ParamDetal;
             деталь.BiesseData = деталь.generate();
             деталь.set(детали[i].ParamDetal, детали[i].ParamStol);
             детали[i].name = детали[i].name.Replace(".", ",");
             

             int coll =int.Parse(детали[i].coll);

             for (int j = 1; j < coll + 1; j++)
             {
                 for (int x = 1; ; x++)
                 {
                     string nameProgr = repStroka(детали[i].name) + "(" + x.ToString() + ")";
                     if (!File.Exists(homeCvet + "\\" + nameProgr + ".fkbpp"))
                     {
                         деталь.Save(homeCvet+"\\" + nameProgr + ".fkbpp");
                         File.WriteAllText(homeCvet + "\\" + nameProgr + ".cix", деталь.generate());
                         break;
                     }
                 }
             }
         }
             return true;
     }

      public static bool getPL(detal деталь)
      {
          if (double.Parse(деталь.DX) * double.Parse(деталь.DY) < 80001f) return true;
          return false;
      }

     public static void CreateProject(string patch,string outPatch,waitForm wait,Form1 form)
     {
         loadThread = new Thread(new ParameterizedThreadStart(Create));
         Alfavit();
         loadThread.Start(new object[4]{patch, outPatch,wait,form});  
     }

     static void closeExel(System.Diagnostics.Process[] oldProc)
     {
         System.Diagnostics.Process[] tekProc = System.Diagnostics.Process.GetProcessesByName("EXCEL");
         for (int i = 0; i < tekProc.Length; i++)
         {
             for (int j = 0; j < oldProc.Length; j++)
             {
                 if (tekProc[i].HandleCount == oldProc[j].HandleCount)
                     break;
                 if (j == oldProc.Length - 1) 
                     tekProc[i].Kill();
             }
         }
     }

     static void Create(object value)
     {
         waitForm wait = (waitForm)((object[])value)[2];
         var tek = new Excel.Application();
         //System.Diagnostics.Process[] oldProc = System.Diagnostics.Process.GetProcessesByName("EXCEL");

         Excel.Workbooks books = tek.Workbooks;
         Excel.Workbook book = books.Open((string)((object[])value)[0]);
         Excel.Worksheet sheet = null;
         

       

         for (int j = 1; j < book.Sheets.Count; j++)
         {
             sheet = book.Sheets[j];
             if (sheet.Name == "Черновик") break;
         }
         if (sheet == null)
         {
             book.Close();
             //closeExel(oldProc);
             wait.Close();
             return;
         }


         string nameZakaz = repRus(sheet.Cells[1, 2].Text);
         if (nameZakaz == "") nameZakaz = DateTime.Now.ToString();
         nameZakaz = nameZakaz.Replace(":", "-").Replace("/", "_").Replace("\\", "_").Replace(".", ",");
         string HomePatch = (string)((object[])value)[1] + nameZakaz;


         if (Directory.Exists(HomePatch))
         {
             if (MessageBox.Show("Такой проект уже существует!\n Пересобрать проект?", "Важное сообщение", MessageBoxButtons.YesNo) == DialogResult.Yes)
                 try
                 {
                     Directory.Delete(HomePatch, true);
                 }
                 catch
                 {
                     for (; ; )
                     {
                         if (MessageBox.Show("Невозможно пересобрать, файлы уже используются в другом приложении!\n Попробовать снова?", "Важное сообщение", MessageBoxButtons.YesNo) == DialogResult.Yes)
                         {
                             try
                             {
                                 Directory.Delete(HomePatch, true);
                                 break;
                             }
                             catch { }
                         }
                         else
                         {
                             wait.progress = 100;
                             return;
                         }

                     }
                 }
             else
             {
                 wait.progress = 100;
                 return;
             }
         }


         ArrayList spisok = new ArrayList();

         string dsp18, dsp16, dvp, dsp10;
         dsp10 = "дсп 10 Белое(кромка Белая0.5)";
         string[] temp = ((string)(sheet.Cells[1, 3].Text)).Split(new string[1]{"#1"},System.StringSplitOptions.None);
         dsp18 = "дсп 18 " +temp[0] + "(кромка "+ temp[1] +")";
         dsp16 = "дсп 16 " + temp[2] + "(кромка " + temp[3] + ")";
         dvp = "двп " + temp[4];


         int collStolov = 0;
         for (int i = 2; i < 3000; i++)
         {
             if (sheet.Cells[i, 1].Text != "" && sheet.Cells[i - 1, 1].Text == "")
             collStolov++;

             if (sheet.Cells[i, 1].Text == "" && sheet.Cells[i + 1, 1].Text == "") break;
         }

         int step = 0;
         if(collStolov > 0)
         step = (int)Math.Round((float)(100 / (collStolov) * 1.8f));


         #region Доп

         stol newStolDsp = new stol() { };
         newStolDsp.name = "Дополнительно";
         newStolDsp.детали = new ArrayList();


         #region Доп дсп
         if (sheet.Cells[1, 15].Text != "")
             {
                 string[] tempFullDsp = ((string)(sheet.Cells[1, 15].Text)).Split(new string[1] {"#2"},System.StringSplitOptions.RemoveEmptyEntries);

                 for (int i = 0; i < tempFullDsp.Length; i++)
                 {
                     string[] tempDsp = tempFullDsp[i].Split('&');
                     string[] cvetStola = (repStroka(tempDsp[8])).Split(new string[1] { "#1" }, System.StringSplitOptions.None);

                     int tekColl = int.Parse(tempDsp[1]);

                     string cvet = "";
                     tempDsp[0] = tempDsp[0].Trim().Replace(",", ".");


                     if (cvetStola.Length < 5)
                         switch (tempDsp[0])
                         {
                             case "18":
                                 cvet = dsp18;
                                 break;
                             case "16":
                                 cvet = dsp16;
                                 break;
                             case "3.2":
                                 cvet = dvp;
                                 break;
                             case "10":
                                 cvet = dsp10;
                                 break;
                         }
                     else
                     {

                         switch (tempDsp[0])
                         {
                             case "18":
                                 cvet = "дсп 18 " + cvetStola[0] + "(кромка " + cvetStola[1] + ")";
                                 break;
                             case "16":
                                 cvet = "дсп 16 " + cvetStola[2] + "(кромка " + cvetStola[3] + ")";
                                 break;
                             case "3.2":
                                 cvet = "двп " + cvetStola[4];
                                 break;
                             case "10":
                                 cvet = dsp10;
                                 break;
                         }
                     }

                     newStolDsp.детали.Add(new detal()
                     {
                       cvet = repRus(repStroka(cvet)),
                       name = tempDsp[2] + " " + tempDsp[4] + "х" + tempDsp[6],
                       DX = tempDsp[4],
                       DY = tempDsp[6],
                       KDX = tempDsp[3],
                       KDY = tempDsp[5],
                       DZ = tempDsp[0],
                       coll = tempDsp[1],
                       Programma = ""
                     });
                 }
             }

         #endregion

         #region Доп фасады
         if (sheet.Cells[1, 145].Text != "")
         {
             string[] tempFullFasad = ((string)(sheet.Cells[1, 145].Text)).Split(new string[1] { "#2" }, System.StringSplitOptions.RemoveEmptyEntries);


             for (int i = 0; i < tempFullFasad.Length; i++)
             {
                 string[] tempDsp = tempFullFasad[i].Split('&');
                 string[] cvetStola = (repStroka(tempDsp[4])).Split(new string[1] { "#1" }, System.StringSplitOptions.None);

                 //int tekColl = int.Parse(tempDsp[2]);

                 string cvet = "";

                 if (tempDsp[4] != "") cvet = "Фасад " + cvetStola[0] + "(кромка " + cvetStola[2] + ")";

                 newStolDsp.детали.Add(new detal()
                 {
                     cvet =repRus(repStroka(cvet)),
                     name = "Фасад "+ tempDsp[0] + " " + tempDsp[1] + "х" + "18",
                     DX = tempDsp[0],
                     DY = tempDsp[1],
                     KDX = "2",
                     KDY = "2",
                     DZ = "18",
                     coll = tempDsp[2],
                     Programma = ""
                 });
             }
         
         
         
         
         
         
         
         }
         #endregion

         if (newStolDsp.детали.Count > 0) spisok.Add(newStolDsp);
         
         #endregion

         #region сбор

         if (step > 0)
         {
             for (int i = 2; i < 3000; i++)
             {
                 if (sheet.Cells[i, 1].Text == "" && sheet.Cells[i + 1, 1].Text == "") break;
                 if (sheet.Cells[i, 1].Text != "" && sheet.Cells[i - 1, 1].Text == "")
                 {
                     int coll = int.Parse(sheet.Cells[i, 7].Text);
                     stol newStol = new stol() { name = repStroka(sheet.Cells[i, 6].Text) };
                     newStol.name = repStroka(newStol.name);
                     newStol.детали = new ArrayList();
                     string[] cvetStola = (repStroka((string)(sheet.Cells[i, 150].Text))).Split(new string[1] { "#1" }, System.StringSplitOptions.None);

                     for (int j = i + 1; j < i + 200; j++)
                     {
                         if (sheet.Cells[j, 1].Text == "") break;
                         if (sheet.Cells[j, 2].Text != "-R" && sheet.Cells[j, 2].Text != "R" && sheet.Cells[j, 3].Text != "0" && sheet.Cells[j, 3].Text != "")
                         {
                             int tekColl = int.Parse(sheet.Cells[j, 3].Text);


                             if (sheet.Cells[j, 2].Text == "f")
                             {
                                 if (sheet.Cells[j, 5].Text == "")
                                 {
                                     continue;
                                     //MessageBox.Show("Не указан цвет фасада в " + newStol.name);
                                     //book.Close();
                                     //closeExel(oldProc);
                                     //wait.Close();
                                     //return;
                                 }
                                 string[] tempCvet = (repStroka((string)sheet.Cells[j, 5].Text)).Split(new string[] { "#1" }, StringSplitOptions.None);
                                 string cvet = "Фасад " + tempCvet[0] + "(кромка " + tempCvet[2] + ")";
                                 newStol.детали.Add(new detal()
                                 {
                                     cvet = repRus(repStroka(cvet)),
                                     DX = sheet.Cells[j, 6].Text,
                                     DY = sheet.Cells[j, 8].Text,
                                     KDX = sheet.Cells[j, 5].Text,
                                     KDY = sheet.Cells[j, 7].Text,
                                     DZ = "18",
                                     name = "Фасад " + sheet.Cells[j, 6].Text + "х" + sheet.Cells[j, 8].Text,
                                     coll = sheet.Cells[j, 3].Text,
                                     Programma = sheet.Cells[j, 13].Text,
                                     ParamDetal = sheet.Cells[j, 14].Text,
                                     ParamStol = sheet.Cells[i, 14].Text
                                 });

                             }
                             else if (sheet.Cells[j, 2].Text != "R")
                             {
                                 string cvet = "";

                                 if (cvetStola.Length < 5)
                                     switch ((string)sheet.Cells[j, 2].Text)
                                     {
                                         case "18":
                                             cvet = dsp18;
                                             break;
                                         case "16":
                                             cvet = dsp16;
                                             break;
                                         case "3.2":
                                             cvet = dvp;
                                             break;
                                         case "10":
                                             cvet = dsp10;
                                             break;
                                     }
                                 else
                                 {

                                     switch ((string)sheet.Cells[j, 2].Text)
                                     {
                                         case "18":
                                             cvet = "дсп 18 " + cvetStola[0] + "(кромка " + cvetStola[1] + ")";
                                             break;
                                         case "16":
                                             cvet = "дсп 16 " + cvetStola[2] + "(кромка " + cvetStola[3] + ")";
                                             break;
                                         case "3.2":
                                             cvet = "двп " + cvetStola[4];
                                             break;
                                         case "10":
                                             cvet = dsp10;
                                             break;
                                     }
                                 }



                                 newStol.детали.Add(new detal()
                                 {
                                     cvet = repRus(repStroka(cvet)),
                                     name = sheet.Cells[j, 4].Text + " " + sheet.Cells[j, 6].Text + "х" + sheet.Cells[j, 8].Text,
                                     DX = sheet.Cells[j, 6].Text,
                                     DY = sheet.Cells[j, 8].Text,
                                     KDX = sheet.Cells[j, 5].Text,
                                     KDY = sheet.Cells[j, 7].Text,
                                     DZ = sheet.Cells[j, 2].Text,
                                     coll = sheet.Cells[j, 3].Text,
                                     Programma = sheet.Cells[j, 13].Text,
                                     ParamDetal = sheet.Cells[j, 14].Text,
                                     ParamStol = sheet.Cells[i, 14].Text
                                 });

                             }

                         }
                     }

                     for (int j = 1; j < coll + 1; j++)
                     {
                         spisok.Add(newStol);
                     }

                     wait.progress += step;
                 }

             }
         }

#endregion

         book.Close();


         //closeExel(oldProc);
         (((Form1)((object[])value)[3])).setLabel(HomePatch + "\\");

         Directory.CreateDirectory(HomePatch);
         HomePatch +="\\";

         for (int i = 0; i < spisok.Count; i++)
         {
             stol _stol = (stol)spisok[i];
             string HomeStol = HomePatch + _stol.name + " (1)";
             if (Directory.Exists(HomeStol))
             {
                 for (int x = 2; ; x++)
                 {
                    HomeStol = HomePatch + _stol.name + " ("+ x.ToString() +")";
                    if (!Directory.Exists(HomeStol))
                    {
                        break;
                    }
                 }
             }

             Directory.CreateDirectory(HomeStol);
             HomeStol += "\\";

             if (!convertShablon(HomeStol, _stol,(Form1)((object[])value)[3]))
             {
                 MessageBox.Show("Шаблон не найден(" + _stol.name +")");
                 //wait.Close();
                 Directory.Delete(HomePatch,true);
                 return;
             };
         }

         try
         {
             wait.progress = 100;
         }
         catch { }

         MessageBox.Show("Готово");

         Form1 form1 = (Form1)((object[])value)[3];
         form1.result = true;
     }

     static string repStroka(string value)
     {
         value = value.Replace("\\", "_").Replace("*", "_");
         value = value.Replace('/', '_').Replace('.', ',').Replace("\"", "");
         return value;
     }

     static string repRus(string value)
     {
         for (int i = 0; i < RusEng.Length; i++)
         {
             string[] temp = RusEng[i].Split('=');
             value = value.Replace(temp[0], temp[1]);
         }
             return value;
     }

    }

    public struct detal
    {
        public string DX, DY, DZ;
        public string KDX, KDY;
        public string Programma;
        public string ParamDetal;
        public string ParamStol;
        public string cvet;
        public string coll;
        public string name;
    }

    public struct stol
    {
        public string name;
        public string coll;
        public ArrayList детали;
    }


    

}
