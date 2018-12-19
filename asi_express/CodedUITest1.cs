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
using Microsoft.VisualStudio.TestTools.UITesting.WpfControls;
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
                            Temp + DirectoryName + @"\express_" + DateTime.Now.ToString("DDMMYYYY_HHmm") + ".log");
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

            File.AppendAllText(Temp + DirectoryName + logfile, tab + txt + "\r\n");
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
            ZipFile.CreateFromDirectory(Temp + DirectoryName, Temp + DirectoryName + this.TestContext.Properties["AgentName"].ToString()+ DateTime.Now.ToString("ddMMyyyy_HHmm") + this.TestContext.CurrentTestOutcome.ToString() + ".zip");
            Directory.Delete(Temp + DirectoryName, true); // удаляем старые данные, так как они нам не нужны больше

        }


        public void Shutdown_asi()
        {
            Process[] p1 = Process.GetProcessesByName("P5");
            foreach(Process Proc in p1)
            {
                Process[] p2 = Process.GetProcessesByName("ASIBusyIndicator_vas_" + Proc.Handle);
                foreach(Process ProcBusy in p2)
                {
                    ProcBusy.Kill();
                }
                Proc.Kill();
            }
            p1 = Process.GetProcessesByName("WINWORD");
            foreach (Process Proc in p1)
            { 
                Proc.Kill();
            }
        }

        public void Clear_cache()
        {
            if (Directory.Exists(LocalAppData + @"\JSC Prognoz"))
            {
                Directory.Delete(LocalAppData + @"\JSC Prognoz", true);
            }
        }

        public void Start_prognoz(int Wac, int lvl)
        {
            if (this.TestContext.Properties["AgentName"].ToString() != "ASI-TST-12-2")
            {
                InputLog(this.TestContext.Properties["AgentName"].ToString(), lvl);
                Process.Start(@"C:\Program Files\JSC Prognoz\Prognoz 5.26\P5.exe");
                InputLog("Открываем прогноз x64", lvl);
            }
            else
            {
                InputLog(this.TestContext.Properties["AgentName"].ToString(), lvl);
                Process.Start(@"C:\Program Files (x86)\JSC Prognoz\Prognoz 5.26\P5.exe");
                InputLog("Открываем прогноз x86", lvl);
            }

        }

        public void StartASI(int WaC, int lvl,string SchemaName, string User)
        {
            this.UIMap.StudioConnectWindow.WaitForControlExist(4 * WaC);
            if (this.TestContext.Properties["AgentName"].ToString() == "ASI-TST-12-2")
            {
                InputLog("Выберем схему "+ SchemaName , lvl);
                this.UIMap.StudioConnectWindow.SchemaConnectWindow.SchemaConnectComboBox.SelectedItem = SchemaName + " @ ASITST11";
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
                this.UIMap.StudioConnectWindow.SchemaConnectWindow.SchemaConnectComboBox.SelectedItem = SchemaName + " @ asi-tst-ms12\\MSSQLSERVER2012";
                InputLog("Введём  логин", lvl);
                this.UIMap.StudioConnectWindow.LoginWindow.LoginEdit.Text = User;
                InputLog("Введём  Пароль", lvl);
                this.UIMap.StudioConnectWindow.PasswordWindow.PasswordEdit.Text = User;
                InputLog("Подтвердим выбор", lvl);
                Mouse.Click(this.UIMap.StudioConnectWindow.OKWindow.OKButton);
            }
        }

        public void StatsArm(int WaC, int lvl, string SchemaName)
        {
            InputLog("Ждём открытия АРМ Админа на стационарном уровне", lvl);
            this.UIMap.ARM_AdminWindow.WaitForControlExist(60 * WaC);
            InputLog("Выберем ТУ", lvl);
            SelectTU(WaC, lvl + 1);
            SetAFSB(WaC, lvl + 1);
            CreateUser(WaC, lvl + 1, this.UIMap.EmployeeWindow.EmployeeList.RIO_3ListItem, "RIO_3", 0, SchemaName);
            DeleteUser(WaC, lvl + 1, this.UIMap.ARM_AdminWindow.MasterWindow.UserPanel.UserList.RIO_3ListItem);
            CreateUser(WaC, lvl + 1, this.UIMap.EmployeeWindow.EmployeeList.RIO_4ListItem, "RIO_4", 1, SchemaName);
            DeleteUser(WaC, lvl + 1, this.UIMap.ARM_AdminWindow.MasterWindow.UserPanel.UserList.RIO_3ListItem);
            RecreateObjects(WaC, lvl);
            this.UIMap.ARM_AdminWindow.SetFocus();
            Keyboard.SendKeys("{F4}", ModifierKeys.Alt);
        }

        public void DeleteUser(int WaC, int lvl, WpfListItem listitem)
        {
            if (listitem.Exists)
            {
                Mouse.Click(listitem);
                Mouse.Click(this.UIMap.ARM_AdminWindow.MasterWindow.UserPanel.DeleteUserButton);
                this.UIMap.AcceptWindow.WaitForControlExist(10 * WaC);
                Mouse.Click(this.UIMap.AcceptWindow.Acc_YesWindow.YesButton);
                this.UIMap.InfoWindow.WaitForControlExist(10 * WaC);
                Mouse.Click(this.UIMap.InfoWindow.OKWindow2.OKButton);
            }
        }
        public void RecreateObjects(int WaC, int lvl)
        {
            Mouse.Click(this.UIMap.ARM_AdminWindow.MasterWindow.NavigationPanel.PrepareButton);
            this.UIMap.ARM_AdminWindow.RecreateButton.WaitForControlExist(15 * WaC);
            Mouse.Click(this.UIMap.ARM_AdminWindow.RecreateButton);

            if(!this.UIMap.ARM_AdminWindow.MasterWindow.RecreateObjectsPanel.RecreateViewCheckBox.Checked)
            {
                this.UIMap.ARM_AdminWindow.MasterWindow.RecreateObjectsPanel.RecreateViewCheckBox.Checked = true;
            }

            if (!this.UIMap.ARM_AdminWindow.MasterWindow.RecreateObjectsPanel.RecreateProcedureCheckBox.Checked)
            {
                this.UIMap.ARM_AdminWindow.MasterWindow.RecreateObjectsPanel.RecreateProcedureCheckBox.Checked = true;
            }

            if (!this.UIMap.ARM_AdminWindow.MasterWindow.RecreateObjectsPanel.AddRigthsCheckBox.Checked)
            {
                this.UIMap.ARM_AdminWindow.MasterWindow.RecreateObjectsPanel.AddRigthsCheckBox.Checked = true;
            }

            Mouse.Click(this.UIMap.ARM_AdminWindow.MasterWindow.RecreateObjectsPanel.RecreateButton);

            this.UIMap.OKWindow.WaitForControlExist(15 * WaC);

            Mouse.Click(this.UIMap.OKWindow.OKButton);
        }
        public void CreateUser(int WaC, int lvl, WpfListItem listitem, string UserName,int Try,string SchemaName) // Try - порядок запуска, если при 0 запуске вводятся данные АИБа, то при последующих уже нет
        {

            Mouse.Click(this.UIMap.ARM_AdminWindow.MasterWindow.NavigationPanel.UsersButton);
            this.UIMap.ARM_AdminWindow.AddUserButton.WaitForControlExist(10 * WaC);
            Mouse.Click(this.UIMap.ARM_AdminWindow.AddUserButton);
            this.UIMap.ARM_AdminWindow.MasterWindow.EditUserPanel.SelectPersonButton.WaitForControlExist(10 * WaC);
            Mouse.Click(this.UIMap.ARM_AdminWindow.MasterWindow.EditUserPanel.SelectPersonButton);
            this.UIMap.EmployeeWindow.WaitForControlExist(10 * WaC);
            if (!listitem.Exists)
            {
                Mouse.Click(this.UIMap.EmployeeWindow.AddEmployeeButton);
                this.UIMap.EmployeeWindow.FIOEmployeeEdit.Text = UserName;
                Mouse.Click(this.UIMap.EmployeeWindow.AddUserButton);
                Mouse.Click(listitem);
            }
            Mouse.Click(this.UIMap.EmployeeWindow.SelectButton);
            this.UIMap.ARM_AdminWindow.MasterWindow.EditUserPanel.LoginEdit.Text = UserName;
            this.UIMap.ARM_AdminWindow.MasterWindow.EditUserPanel.PasswordEdit.Text = UserName;
            this.UIMap.ARM_AdminWindow.MasterWindow.EditUserPanel.RePasswordEdit.Text = UserName;
            if (!this.UIMap.ARM_AdminWindow.MasterWindow.EditUserPanel.AsiAutoStart.Selected)
            {
                this.UIMap.ARM_AdminWindow.MasterWindow.EditUserPanel.AsiAutoStart.Selected = true;
            }
            Mouse.Click(this.UIMap.ARM_AdminWindow.MasterWindow.EditUserPanel.SaveUserButton);   
            if (Try == 0)
            {
                this.UIMap.ISA_Window.WaitForControlExist(30 * WaC);
                this.UIMap.ISA_Window.IsaLoginEdit.Text = SchemaName + "_ISA";
                this.UIMap.ISA_Window.IsaPasswordEdit.Text = SchemaName + "_ISA";
                Mouse.Click(this.UIMap.ISA_Window.OKButton);
            }
            this.UIMap.UserCreatedWindow.OKWindow.OKButton.WaitForControlExist(30 * WaC);
            Mouse.Click(this.UIMap.UserCreatedWindow.OKWindow.OKButton);
        }
        public void SelectTU(int WaC, int lvl)
        {
            InputLog("Раскроем панель подготовки", lvl);
            Mouse.Click(this.UIMap.ARM_AdminWindow.MasterWindow.NavigationPanel.PrepareButton);
            InputLog("Выберем регион базирования", lvl);
            Mouse.Click(this.UIMap.ARM_AdminWindow.TuPropButton);
            InputLog("Ждём появления комбобокса с регионами", lvl);
            this.UIMap.ARM_AdminWindow.MasterWindow.TuPropWindow.RegionsComboBox.WaitForControlExist(15 * WaC);
            Mouse.Click(this.UIMap.ARM_AdminWindow.MasterWindow.TuPropWindow.RegionsComboBox);
            InputLog("Ткнём в него", lvl);
            this.UIMap.ARM_AdminWindow.Tree.TreeLvL1.WaitForControlExist(15 * WaC);
            InputLog("Ткнём в первый уровень дерева элементов", lvl);
            Mouse.Click(this.UIMap.ARM_AdminWindow.Tree.TreeLvL1);
            InputLog("Спустимся ниже", lvl);
            WaiterForMSSQL(2);
            Keyboard.SendKeys("{DOWN}");
            InputLog("Раскроем второй уровень дерева элементов", lvl);
            WaiterForMSSQL(2);
            Keyboard.SendKeys("{RIGHT}");
            InputLog("Дождёмся его появления", lvl);
            this.UIMap.ARM_AdminWindow.Tree.TreeLvL1.TreeLvL2.WaitForControlExist(1 * WaC);
            InputLog("В зависимости от выбранного элемента выберем 117 или 118 ID ", lvl);
            WaiterForMSSQL(2);
            if (this.UIMap.ARM_AdminWindow.MasterWindow.TuPropWindow.RegionsComboBox.SelectedIndex == 118 ||
                    this.UIMap.ARM_AdminWindow.MasterWindow.TuPropWindow.RegionsComboBox.SelectedIndex == -1)
            {
                InputLog("Спустимся ниже", lvl);
                WaiterForMSSQL(2);
                Keyboard.SendKeys("{DOWN}");
            }
            else
            {
                InputLog("Спустимся ниже", lvl);
                WaiterForMSSQL(2);
                Keyboard.SendKeys("{DOWN}");
                InputLog("Спустимся ниже", lvl);
                WaiterForMSSQL(2);
                Keyboard.SendKeys("{DOWN}");
            }
            InputLog("Подтвердим свой выбор", lvl);
            Keyboard.SendKeys("{ENTER}");
            InputLog("Дождёмся появления текстовой надписи на форме с названием региона", lvl);
            this.UIMap.ARM_AdminWindow.MasterWindow.TuPropWindow.NameTextBlock.WaitForControlExist(1500 * WaC);
            InputLog("Дождёмся исчезновния дерева элементов", lvl);
            this.UIMap.ARM_AdminWindow.Tree.WaitForControlNotExist(10 * WaC);
        }

        private void WaiterForMSSQL(int Time)
        {
            if (this.TestContext.Properties["AgentName"].ToString() != "ASI-TST-12-2") { Thread.Sleep(Time*1000); }
        }

        public void SetAFSB(int WaC, int lvl)
        {
            InputLog("Нажмём кнопку подготовки системы", lvl);
            Mouse.Click(this.UIMap.ARM_AdminWindow.MasterWindow.NavigationPanel.PrepareButton);
            InputLog("Нажмём кнопку настройки соединения с САФСБ", lvl);
            Mouse.Click(this.UIMap.ARM_AdminWindow.SAFSBButton);
            InputLog("Ждём появления элементов для ввода данных", lvl);
            this.UIMap.ARM_AdminWindow.SAFSBData.WaitForControlExist(10 * WaC);
            if (this.TestContext.Properties["AgentName"].ToString() == "ASI-TST-12-2")
            {
                InputLog("Ввёдем TCP", lvl);
                this.UIMap.ARM_AdminWindow.SAFSBData.ServerProperties.ProtocolEdit.Text = "TCP";
                InputLog("Ввёдем HOST", lvl);
                this.UIMap.ARM_AdminWindow.SAFSBData.ServerProperties.HostEdit.Text = "10.7.0.33";
                InputLog("Ввёдем Port", lvl);
                this.UIMap.ARM_AdminWindow.SAFSBData.ServerProperties.PortEdit.Text = "1521";
                InputLog("Ввёдем SID", lvl);
                this.UIMap.ARM_AdminWindow.SAFSBData.ServerProperties.ServiceEdit.Text = "ASITST11";
                InputLog("Ввёдем название схемы тех.пользователя", lvl);
                this.UIMap.ARM_AdminWindow.TechUserData.TechUserDataGroup.TechUserLoginEdit.Text = "ASI_TECHUSER_UI_" + DateTime.Now.ToString("ddMMyyyy_HHmm");
                InputLog("Ввёдем пароль схемы тех.пользователя", lvl);
                this.UIMap.ARM_AdminWindow.TechUserData.TechUserDataGroup.TechUserPasswordEdit.Text = "ASI_TECHUSER_UI_" + DateTime.Now.ToString("ddMMyyyy_HHmm");
            }
            else
            {
                InputLog("Сбросим данные в IP сервера", lvl);
                this.UIMap.ARM_AdminWindow.SAFSBData.ServerMSSQLEdit.Text = "";
                InputLog("Сбросим данные в названии схемы", lvl);
                this.UIMap.ARM_AdminWindow.SAFSBData.SchemaNameEdit.Text = "";
                InputLog("Ввёдем данные IP сервера", lvl);
                Keyboard.SendKeys(this.UIMap.ARM_AdminWindow.SAFSBData.ServerMSSQLEdit, "10.7.0.32");
                InputLog("Ввёдем данные названия схемы", lvl);
                Keyboard.SendKeys(this.UIMap.ARM_AdminWindow.SAFSBData.SchemaNameEdit, "RDATU71_DATA");
                InputLog("Ввёдем данные логина тех.пользователя", lvl);
                this.UIMap.ARM_AdminWindow.TechUserData.TechUserDataGroup.TechUserLoginEdit.Text = "ASI_TECHUSER_UI_" + DateTime.Now.ToString("ddMMyyyy_HHmm");
                InputLog("Ввёдем данные пароля тех.пользователя", lvl);
                this.UIMap.ARM_AdminWindow.TechUserData.TechUserDataGroup.TechUserPasswordEdit.Text = "Qwerty1";
            }
            InputLog("Проверим выбран ли комбобокс на создания тех.пользователя", lvl);
            if (!this.UIMap.ARM_AdminWindow.TechUserData.NeedCreateTechUserCheckBox.Checked)
            {
                InputLog("Выберем его", lvl);
                Mouse.Click(this.UIMap.ARM_AdminWindow.TechUserData.NeedCreateTechUserCheckBox);
                InputLog("Ждём появления поля для ввода системный параметров", lvl);
                this.UIMap.ARM_AdminWindow.TechUserData.SYSTEMGroup.SystemLoginEdit.WaitForControlExist(1 * WaC);
                InputLog("Ввёдем логин суперпользователя", lvl);
                this.UIMap.ARM_AdminWindow.TechUserData.SYSTEMGroup.SystemLoginEdit.Text = this.TestContext.Properties["SysLogin"].ToString();
                InputLog("Ввёдем пароль суперпользователя", lvl);
                this.UIMap.ARM_AdminWindow.TechUserData.SYSTEMGroup.SystemPasswordEdit.Text = this.TestContext.Properties["SysPassword"].ToString();
            }
            InputLog("Создадим подключение", lvl);
            Mouse.Click(this.UIMap.ARM_AdminWindow.MasterWindow.SafsbPanel.CreateConnectSafsb);
            InputLog("Ожидаем завершения создания подключения", lvl);
            this.UIMap.OKWindow.OKButton.WaitForControlExist(120 * WaC);
            InputLog("Проверим, раскрыто ли окно с логами", lvl);
            if (!this.UIMap.ARM_AdminWindow.TechUserData.LogWindow.Expanded)
            {
                InputLog("Раскроем", lvl+1);
                this.UIMap.ARM_AdminWindow.TechUserData.LogWindow.Expanded = true;
            }
            InputLog("Проверим, есть ли там ошибки", lvl);
            if (this.UIMap.ARM_AdminWindow.TechUserData.LogWindow.TextLog.Text.Contains("Ошибка:"))
            {
                InputLog("Ошибки есть", lvl +1);
                InputLog(this.UIMap.ARM_AdminWindow.TechUserData.LogWindow.TextLog.Text, lvl + 1);
                this.UIMap.ARM_AdminWindow.TechUserData.LogWindow.TextLog.WaitForControlNotExist(1 * WaC);
                GetScreen("Error_after_set_settings_AFSB" + this.TestContext.Properties["AgentName"].ToString());
                // введено специально, для того чтобы  отловить ошибки при формировании
            }
        }

        public void MobArm(int WaC, int lvl, string SchemaName)
        {
            InputLog("Ждём открытия АРМ Админа на мобильном уровне", lvl);
            this.UIMap.ARM_AdminWindow.WaitForControlExist(30 * WaC);
        }


        public void PrepareAsi(int WaC, int lvl, string SchemaName)
        {
            Clear_cache();
            InputLog("Запустим прогноз", lvl);
            Start_prognoz(WaC, lvl + 1);
            InputLog("Подключимся под Администратором и проверим/восстановим все настройки", lvl);
            StartASI(WaC, lvl + 1, SchemaName, SchemaName); 
            if (SchemaName == "ASISTA_UI")
            {
                StatsArm(WaC, lvl + 1, SchemaName);
            }
            else
            {
                MobArm(WaC, lvl + 1, SchemaName);
            }
        }


        #endregion


        [TestMethod]
        [TestProperty("AgentName", "ASI-TST-MS12")]
        [TestProperty("Files", "FileToDeploy.txt")]
        [TestProperty("SysLogin", "sa")]
        [TestProperty("SysPassword", "Qwerty1")]
        public void Asi_Express_MSSQL() => Asi_Express_All(1000);

        [TestMethod]
        [TestProperty("AgentName", "ASI-TST-12-2")]
        [TestProperty("Files", "FileToDeploy.txt")]
        [TestProperty("SysLogin", "SYSTEM")]
        [TestProperty("SysPassword", "ASITST11")]
        public void Asi_Express_ORACLE() => Asi_Express_All(4000);



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
