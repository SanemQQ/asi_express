using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;
using System.Diagnostics;
using System.IO;
using System.Globalization;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;


namespace asi_express
{

    //    Dictionary Dict 

    [CodedUITest]
    public class asi_express
    {


        #region Additional test attributes
        string LocalAppData = Environment.GetEnvironmentVariable("LocalAppData");
        string Temp = Environment.GetEnvironmentVariable("Temp");
        string DirectoryName = @"\asi_express_log";
        CultureInfo lang = new CultureInfo("ru-RU");
        Dictionary<string, Point> Dots = new Dictionary<string, Point>();

        [TestInitialize]
        public void TestStartup()
        {
            if (!Directory.Exists(Temp + DirectoryName))
            {
                Directory.CreateDirectory(Temp + DirectoryName);
            }
            if (!File.Exists(Temp+DirectoryName+@"\express_"+this.TestContext.Properties["AgentName"].ToString()+".log"))
            {
                File.AppendAllText(Temp + DirectoryName + @"\express_" + this.TestContext.Properties["AgentName"].ToString() + ".log","");
            }
            else
            {
                File.AppendAllText(Temp + DirectoryName + @"\express_" + this.TestContext.Properties["AgentName"].ToString() + DateTime.UtcNow.ToShortDateString().ToString(lang) + ".log", "");
                File.Copy(Temp + DirectoryName + @"\express_" + this.TestContext.Properties["AgentName"].ToString() + ".log",
                            Temp + DirectoryName + @"\express_" + this.TestContext.Properties["AgentName"].ToString() + DateTime.UtcNow.ToShortDateString().ToString(lang) + ".log");
                File.AppendAllText(Temp + DirectoryName + @"\express_" + this.TestContext.Properties["AgentName"].ToString() + ".log", "");                
            }

            Dots.Add("", new Point());
            

        }

        [TestCleanup]
        public void MyTestCleanup()
        {
           // Console.WriteLine(@"\System32\calc.exe");
        }
        #endregion


        [TestMethod]
        public void Asi_Express_MSSQL ()
        {
            Asi_Express_All(5);
        }

        [TestMethod]
        public void Asi_Express_ORACLE()
        {
            Asi_Express_All(15);
        }





        public void Asi_Express_All(int WaC) // WaC - Waiter Coefficient - коэффициент ожидания, который будет корректировать время ожидания между ораклом и мсскл
        {
            Process.Start(WinDir+@"\System32\calc.exe");
        }



        public void inputLog (string txt)
        {


        }




        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }
        private TestContext testContextInstance;

        public UIMap UIMap
        {
            get
            {
                if (this.map == null)
                {
                    this.map = new UIMap();
                }

                return this.map;
            }
        }

        private UIMap map;
    }
}
