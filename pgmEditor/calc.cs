using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.CSharp;
using System.CodeDom.Compiler;
using System.Reflection;

namespace pgmEditor
{
    public  class calculator
    {
         private  string begin = @"using System;
                                        namespace MyNamespace
                                       {
                                         public delegate string Calc();
                                         public static class LambdaCreator 
                                         {
                                            public static Calc Create()
                                           {
                                            return ()=>(";
        private  string end = @").ToString();
                                           }
                                         }
                                       }";

        CSharpCodeProvider provider = new CSharpCodeProvider();
        CompilerParameters parameters = new CompilerParameters();

        public calculator()
        {
            parameters.GenerateInMemory = true;
            parameters.GenerateExecutable = false;
            parameters.ReferencedAssemblies.Add("System.dll");
        }


        public  string Result(string stroka)
        {

            try
            {
                float test = float.Parse(stroka.Replace(".",","));
                return test.ToString();

            }
            catch { }

            try
            {
                if (stroka.Contains("="))
                    if (!stroka.Contains(">") && !stroka.Contains("<"))
                        stroka = stroka.Replace("=", "==");


                stroka = stroka.Replace(",", ".");

                if (stroka.Contains("."))
                {
                    string[] temp = stroka.Split('.');

                    for (int i = 1; i < temp.Length; i++)
                    {
                        Char[] str = temp[i].ToCharArray();
                        string outstr = "";
                        bool nashel = false;
                        for (int j = 0; j < str.Length; j++)
                        {
                            if (!nashel)
                                if (!Char.IsDigit(str[j]))
                                {
                                    outstr += "f";
                                    nashel = true;
                                }
                            outstr += str[j];
                        }

                        temp[i] = outstr;
                    }

                    stroka = String.Join(".", temp);
                }

                CompilerResults results = provider.CompileAssemblyFromSource(parameters, begin + stroka + end);
                var cls = results.CompiledAssembly.GetType("MyNamespace.LambdaCreator");
                var method = cls.GetMethod("Create", BindingFlags.Static | BindingFlags.Public);
                var calc = (method.Invoke(null, null) as Delegate);
                return calc.DynamicInvoke().ToString();
            }
            catch { }
            return null;
        }

    }
}
