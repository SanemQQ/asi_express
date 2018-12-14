using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
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
        readonly string LocalAppData = Environment.GetEnvironmentVariable("LocalAppData");
        readonly string Temp = Environment.GetEnvironmentVariable("Temp");
        readonly string DirectoryName = @"\asi_express_log";
        CultureInfo lang = new CultureInfo("ru-RU");
        Dictionary<string, Point> Dots = new Dictionary<string, Point>();
        readonly string logfile = @"\express.log";

        [TestInitialize]
        public void TestStartup()
        {
            if (!Directory.Exists(Temp + DirectoryName))
            {
                Directory.CreateDirectory(Temp + DirectoryName);
            }
            if (!File.Exists(Temp+DirectoryName+@"\express.log"))
            {
                File.AppendAllText(Temp + DirectoryName + @"\express.log","");
            }
            else
            {
                File.AppendAllText(Temp + DirectoryName + @"\express_"  + DateTime.UtcNow.ToShortDateString().ToString(lang) + ".log", "");
                File.Copy(Temp + DirectoryName + @"\express.log",
                            Temp + DirectoryName + @"\express_" + DateTime.UtcNow.ToShortDateString().ToString(lang) + ".log");
                File.AppendAllText(Temp + DirectoryName + @"\express.log", "");                
            }
            // Справочник точек, по которым будет проводиться взаимодействие с некоторыми объектами
            Dots.Add("", new Point());
            

        }

        public void inputLog( string txt , int lvl )  // txt - строка, которую передают для логирования, а lvl - это уровень вложенности
        {
            string tab = "";
            for (int i = 0; i < lvl; i++)
            {
                tab = tab + "    ";
            }

            File.AppendAllText(Temp + DirectoryName + logfile, tab + txt);
        }

        public void GetScreen(string imgName) 
        {
            using (var image = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height))
            {
                using (var graphics = Graphics.FromImage(image))
                    graphics.CopyFromScreen(0, 0, 0, 0, image.Size);

                image.Save(Temp+DirectoryName + imgName, System.Drawing.Imaging.ImageFormat.Jpeg);
            }
        }

        [TestCleanup]
        public void MyTestCleanup()
        {
            Shutdown_asi();
            Clear_cache();
            ZipFile.CreateFromDirectory(Temp + DirectoryName, Temp + DirectoryName + this.TestContext.Properties["AgentName"].ToString()+this.TestContext.CurrentTestOutcome.ToString() + ".zip");
            Directory.Delete(Temp + DirectoryName, true); // удаляем старые данные, так как они нам не нужны больше

        }


        public void Shutdown_asi()
        {
            Process[] p1 = Process.GetProcessesByName("Studio.exe");
            foreach(Process Proc in p1)
            {
                Process[] p2 = Process.GetProcessesByName("ASIBusyIndicator_vas_" + Proc.Handle + ".exe");
                foreach(Process ProcBusy in p2)
                {
                    ProcBusy.Kill();
                }
                Proc.Kill();
            }
            p1 = Process.GetProcessesByName("WINWORD.exe");
            foreach (Process Proc in p1)
            { 
                Proc.Kill();
            }
        }

        public void Clear_cache()
        {
            Directory.Delete(LocalAppData + @"\JSC Prognoz",true);
        }

        public void start_asi()
        {



        }


        #endregion


        [TestMethod]
        [TestProperty("AgenName","ASI-TST-MS12")]
        [TestProperty("Files","FileToDeploy.txt")]
        public void Asi_Express_MSSQL ()
        {
            Asi_Express_All(5);
        }

        [TestMethod]
        [TestProperty("AgenName", "ASI-TST-12-2")]
        [TestProperty("Files", "FileToDeploy.txt")]
        public void Asi_Express_ORACLE()
        {
            Asi_Express_All(15);
        }





        public void Asi_Express_All(int WaC) // WaC - Waiter Coefficient - коэффициент ожидания, который будет корректировать время ожидания между ораклом и мсскл
        {

            try
            {
            

            }
            catch (Exception e)
            {
                inputLog(e.Message, 0);
                throw e;
            }

        }




        #region UsefulMethods






        #endregion UsefulMethods





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
