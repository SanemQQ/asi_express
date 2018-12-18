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
    // Комментарии к параметрам методов
    // WaC - коэффициент ожидания в зависимости от  оракла/мсскл
    // lvl - вложенность лога в текстовом файле

    [CodedUITest]
    public class Asi_express
    {


        #region Additional test attributes
        readonly string LocalAppData = Environment.GetEnvironmentVariable("LocalAppData");
        readonly string Temp = Environment.GetEnvironmentVariable("Temp");
        readonly string DirectoryName = @"\asi_express_log";
        readonly CultureInfo lang = new CultureInfo("ru-RU");
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

        public void InputLog( string txt , int lvl )  // txt - строка, которую передают для логирования, а lvl - это уровень вложенности
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

        public void Start_prognoz(int Wac, int lvl)
        {
            if (this.TestContext.Properties["AgentName"].ToString() == "ASI-TST-12-2")
            {
                Process.Start(@"C:\Program Files\JSC Prognoz\Prognoz 5.26\P5.exe");
                InputLog("Открываем прогноз x64", lvl);
            }
            else
            {
                Process.Start(@"C:\Program Files (x86)\JSC Prognoz\Prognoz 5.26\P5.exe");
                InputLog("Открываем прогноз x86", lvl);
            }

        }

        public void StartASI(int WaC, int lvl,string SchemaName, string User)
        {
            this.UIMap.StudioConnectWindow.WaitForControlExist(4000 * WaC);
            if (this.TestContext.Properties["AgentName"].ToString() == "ASI-TST-12-2")
            {
                InputLog("Выберем схему "+ SchemaName , lvl);
                this.UIMap.StudioConnectWindow.SchemaConnectWindow.SchemaConnectComboBox.SelectedItem = SchemaName + "@ ASITST11";
                InputLog("Введём  логин", lvl);
                this.UIMap.StudioConnectWindow.LoginWindow.LoginEdit.Text = User;
                InputLog("Введём  Пароль", lvl);
                this.UIMap.StudioConnectWindow.PasswordWindow.PasswordEdit.Text = User;
                InputLog("Подтвердим выбор", lvl);
                Mouse.Click(this.UIMap.StudioConnectWindow.OKWindow.OKButton);
            }
            else
            {
                InputLog("Выберем схему " + SchemaName, lvl);
                this.UIMap.StudioConnectWindow.SchemaConnectWindow.SchemaConnectComboBox.SelectedItem = SchemaName + "@ asi-tst-ms12\\MSSQLSERVER2012";
                InputLog("Введём  логин", lvl);
                this.UIMap.StudioConnectWindow.LoginWindow.LoginEdit.Text = User;
                InputLog("Введём  Пароль", lvl);
                this.UIMap.StudioConnectWindow.PasswordWindow.PasswordEdit.Text = User;
                InputLog("Подтвердим выбор", lvl);
                Mouse.Click(this.UIMap.StudioConnectWindow.OKWindow.OKButton);
            }
        }

        public void StatsArm(int WaC, int lvl)
        {
            InputLog("Ждём открытия АРМ Админа на стационарном уровне", lvl);
            this.UIMap.ARM_AdminWindow.WaitForControlExist(30000 * WaC);
            InputLog("Раскроем панель подготовки", lvl);
            Mouse.Click(this.UIMap.ARM_AdminWindow.MasterWindow.NavigationPanel.PrepareButton);
            InputLog("Выберем регион базирования", lvl);
            Mouse.Click(this.UIMap.ARM_AdminWindow.TuPropButton);
            InputLog(":l`v", lvl);
            this.UIMap.ARM_AdminWindow.MasterWindow.TuPropWindow.RegionsComboBox.WaitForControlExist(15000 * WaC);
            Mouse.Click(this.UIMap.ARM_AdminWindow.MasterWindow.TuPropWindow.RegionsComboBox);
            if (this.UIMap.ARM_AdminWindow.MasterWindow.TuPropWindow.RegionsComboBox.SelectedIndex == 118 ||
                    this.UIMap.ARM_AdminWindow.MasterWindow.TuPropWindow.RegionsComboBox.SelectedIndex == - 1)
            {
                Mouse.Click(this.UIMap.ARM_AdminWindow.Tree.TreeLvL1);
                Keyboard.SendKeys("{DOWN}");
                Keyboard.SendKeys("{RIGTH}");
                this.UIMap.ARM_AdminWindow.Tree.TreeLvL1.TreeLvL2.WaitForControlExist(1000 * WaC);
                Keyboard.SendKeys("{DOWN}");
                Keyboard.SendKeys("{ENTER}");

            }
            else
            {
                Mouse.Click(this.UIMap.ARM_AdminWindow.Tree.TreeLvL1);
                Keyboard.SendKeys("{DOWN}");
                Keyboard.SendKeys("{RIGTH}");
                this.UIMap.ARM_AdminWindow.Tree.TreeLvL1.TreeLvL2.WaitForControlExist(1000 * WaC);
                Keyboard.SendKeys("{DOWN}");
                Keyboard.SendKeys("{DOWN}");
                Keyboard.SendKeys("{ENTER}");
            }

            this.UIMap.ARM_AdminWindow.MasterWindow.TuPropWindow.NameTextBlock.WaitForControlExist(1500 * WaC);


            Mouse.Click(this.UIMap.ARM_AdminWindow.MasterWindow.NavigationPanel.PrepareButton);
            Mouse.Click(this.UIMap.ARM_AdminWindow.SAFSBButton);
            this.UIMap.ARM_AdminWindow.SAFSBData.WaitForControlExist(10000 * WaC);
            if (this.TestContext.Properties["AgentName"].ToString() == "ASI-TST-12-2")
            {

            }
            else
            {
               // this.UIMap.
            }
        }

        public void MobArm(int WaC, int lvl)
        {
            InputLog("Ждём открытия АРМ Админа на мобильном уровне", lvl);
            this.UIMap.ARM_AdminWindow.WaitForControlExist(30000 * WaC);
        }


        public void PrepareAsi(int WaC, int lvl, string SchemaName)
        {
            InputLog("Запустим прогноз", lvl);
            Start_prognoz(WaC, lvl + 1);
            InputLog("Подключимся под Администратором и проверим/восстановим все настройки", lvl);
            StartASI(WaC, lvl + 1, SchemaName, SchemaName); 
            if (SchemaName == "ASISTA_UI")
            {
                StatsArm(WaC, lvl + 1);
            }
            else
            {
                MobArm(WaC, lvl + 1);
            }
        }


        #endregion


        [TestMethod]
        [TestProperty("AgentName","ASI-TST-MS12")]
        [TestProperty("Files","FileToDeploy.txt")]
        public void Asi_Express_MSSQL ()
        {
            Asi_Express_All(2);
        }

        [TestMethod]
        [TestProperty("AgentName", "ASI-TST-12-2")]
        [TestProperty("Files", "FileToDeploy.txt")]
        public void Asi_Express_ORACLE()
        {
            Asi_Express_All(8);
        }





        public void Asi_Express_All(int WaC) // WaC - Waiter Coefficient - коэффициент ожидания, который будет корректировать время ожидания между ораклом и мсскл
        {

            try
            {
                InputLog("Начнём подготовку к экспресс-тестированию", 0);
                PrepareAsi(WaC, 1, "ASISTA_UI");

            }
            catch (Exception e)
            {
                InputLog(e.Message, 0);
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
