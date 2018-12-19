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
                            Temp + DirectoryName + @"\express_" + DateTime.Now.ToString("ddMMyyyy_HHmm") + ".log");
                File.AppendAllText(Temp + DirectoryName + logfile , Temp + DirectoryName + " \r\n " +
                                                                            LocalAppData +"\r\n"+ InputLanguage.CurrentInputLanguage.ToString()
                                                                            +"\r\n" + InputLanguage.DefaultInputLanguage.ToString()
                                                                            + CultureInfo.CurrentCulture.ToString()
                                                                            + CultureInfo.CurrentUICulture.ToString());                
            }
            // Справочник точек, по которым будет проводиться взаимодействие с некоторыми объектами
            //RibbonUpperButtons
            Dots.Add("InfoManag", new Point(305,45));
            //RibbonButtons
            Dots.Add("AFSBButton", new Point(165,80));
            Dots.Add("SPRButton", new Point(55, 80));

            //ImportFromAFSB
            Dots.Add("101F", new Point(235, 290));
            Dots.Add("102F", new Point(235, 320));
            Dots.Add("117F", new Point(235, 440));
            Dots.Add("118F", new Point(235, 470));
            Dots.Add("501F", new Point(235, 830));
            Dots.Add("SPRF", new Point(235, 930));
            Dots.Add("FirstPointScroll", new Point(1865, 290));
            Dots.Add("SecondPointScroll", new Point(1865, 830));
            Dots.Add("UpperCalendar", new Point(195, 220));
            Dots.Add("BottomCalendar", new Point(195, 250));

            //CalcSPR

            Dots.Add("SprAFSBCalc", new Point(65, 250));

            // InputLog(Environment.OSVersion.ToString(), 0);

            //FormSetQuest
            Dots.Add("AddTask", new Point(95, 205));
            Dots.Add("ChangeCode", new Point(445, 230));
            Dots.Add("ChangeEmp", new Point(1230, 230));
            Dots.Add("131Code", new Point(683, 478));
            Dots.Add("FirstEmp", new Point(595, 365));
            Dots.Add("SecondEmp", new Point(595, 380));
            Dots.Add("SaveCode", new Point(50, 80));


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

                image.Save(Temp + DirectoryName + "\\" + imgName + ".jpeg", System.Drawing.Imaging.ImageFormat.Jpeg);
            }
        }

        [TestCleanup]
        public void MyTestCleanup()
        {
            Shutdown_asi();
            Thread.Sleep(10000);
            Clear_cache();
            ZipFile.CreateFromDirectory(Temp + DirectoryName, Temp + DirectoryName + this.TestContext.Properties["AgentName"].ToString()+ DateTime.Now.ToString("ddMMyyyy_HHmm") + this.TestContext.CurrentTestOutcome.ToString() + ".zip");
            Directory.Delete(Temp + DirectoryName, true); // удаляем старые данные, так как они нам не нужны больше
        }

        public void OpenAFSB(int WaC, int lvl)
        {
            InputLog("Раскроем вкладку информационное обеспечение", lvl);
            Mouse.Click(Dots["InfoManag"]);
            Thread.Sleep(3 * WaC);
            InputLog("Нажмём кнопку загрузки данных из АФСБ", lvl);
            Mouse.Click(Dots["AFSBButton"]);
            Thread.Sleep(3 * WaC);
            InputLog("Дождёмся её появления", lvl);
            this.UIMap.ASI_Window.ImportAFSB.WaitForControlExist(60 * WaC);
        }

        public void DownloadSpr(int WaC, int lvl)
        {
            InputLog("Открываем вкладку загрузки АФСБ", lvl);
            OpenAFSB(WaC, lvl+1);
            InputLog("Перешли к первой точке", lvl);
            Mouse.Move(Dots["FirstPointScroll"]);
            InputLog("Тащим скролл", lvl);
            Mouse.StartDragging();
            InputLog("До этой точки", lvl);
            Mouse.StopDragging(Dots["SecondPointScroll"]);
            InputLog("Ждём пока отстроится", lvl);
            Thread.Sleep(2 * WaC);
            InputLog("Выделяем чекбокс справочной информации", lvl);
            Mouse.Click(Dots["SPRF"]);
            InputLog("Жмём кнопку загрузки", lvl);
            Mouse.Click(this.UIMap.ASI_Window.ImportAFSB.ImportWindow.ImportButton);
            InputLog("Ждём подтверждающего окна", lvl);
            this.UIMap.AcceptWindow.Acc_YesWindow.YesButton.WaitForControlExist(20 * WaC);
            InputLog("Соглашаемся", lvl);
            Mouse.Click(this.UIMap.AcceptWindow.Acc_YesWindow.YesButton);
            InputLog("Ждём окно протокола", lvl);
            this.UIMap.LogWindow.WaitForControlExist(3000 * WaC);
            InputLog("Скриним", lvl);
            GetScreen("FinishDownload");
            InputLog("Закрываем", lvl);
            Thread.Sleep(5 * WaC);
            Mouse.Click(this.UIMap.LogWindow.CloseWindow.CloseButton);
        }

        public void CalcSpr(int WaC, int lvl)
        {
            InputLog("Раскроем вкладку информационное обеспечение", lvl);
            Mouse.Click(Dots["InfoManag"]);
            Thread.Sleep(3 * WaC);
            InputLog("Нажмём на кнопку справочники", lvl);
            Mouse.Click(Dots["SPRButton"]);
            InputLog("Дождёмся появления формы расчета справочников", lvl);
            this.UIMap.ASI_Window.CalcSpr.WaitForControlExist(60 * WaC);
            Thread.Sleep(3 * WaC);
            InputLog("Нажмём на чекбокс расчёта справочников на основании АФСБ", lvl);
            Mouse.Click(Dots["SprAFSBCalc"]);
            InputLog("Рассчитаем данные", lvl);
            Mouse.Click(this.UIMap.ASI_Window.CalcSpr.UploadWindow.UploadButton);
            InputLog("Ждём окно с логами", lvl);
            this.UIMap.LogWindow.WaitForControlExist(3000 * WaC);
            InputLog("Скриним", lvl);
            GetScreen("FinishСфдс");
            InputLog("Закрываем", lvl);
            Thread.Sleep(5 * WaC);
            Mouse.Click(this.UIMap.LogWindow.CloseWindow.CloseButton);
            Mouse.Click(Dots["SaveCode"]);
            this.UIMap.AcceptWindow.Acc_YesWindow.YesButton.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.AcceptWindow.Acc_YesWindow.YesButton);
            this.UIMap.AcceptWindow.Acc_YesWindow.YesButton.WaitForControlNotExist(60 * WaC);
            Mouse.Click(this.UIMap.AcceptWindow.Acc_YesWindow.YesButton);
        }

        public void CreateQuestions(int WaC, int lvl)
        {
            Mouse.Click(Dots["AddTask"]);
            Thread.Sleep(1 * WaC);
            Mouse.Click(Dots["ChangeCode"]);
            this.UIMap.SelectCode.WaitForControlExist(60 * WaC);
            Mouse.Click(Dots["131Code"]);
            Thread.Sleep(1 * WaC);
            Mouse.Click(this.UIMap.SelectCode.OKWindow.OKButton);
            this.UIMap.SelectCode.WaitForControlNotExist(60 * WaC);
            Mouse.Click(Dots["ChangeEmp"]);
            this.UIMap.FormSprEmp.WaitForControlExist(60 * WaC);
            Mouse.Click(Dots["FirstEmp"]);
            Thread.Sleep(1 * WaC);
            Mouse.Click(Dots["SecondEmp"]);
            Mouse.Click(this.UIMap.FormSprEmp.OKWindow.OKButton);
            this.UIMap.FormSprEmp.WaitForControlNotExist(60 * WaC);
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
            if (this.TestContext.Properties["AgentName"].ToString() == "ASI-TST-12-2")
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
            if (Environment.OSVersion.ToString() != "Microsoft Windows NT 6.1.7601 Service Pack 1")
            {
          //     SelectTU(WaC, lvl + 1);
            }
          //  SetAFSB(WaC, lvl + 1);
            CreateUser(WaC, lvl + 1, this.UIMap.EmployeeWindow.EmployeeList.RIO_3ListItem, "RIO_3", 0, SchemaName);
            DeleteUser(WaC, lvl + 1, this.UIMap.ARM_AdminWindow.MasterWindow.UserPanel.UserList.RIO_3ListItem);
            CreateUser(WaC, lvl + 1, this.UIMap.EmployeeWindow.EmployeeList.RIO_4ListItem, "RIO_4", 1, SchemaName);
            DeleteUser(WaC, lvl + 1, this.UIMap.ARM_AdminWindow.MasterWindow.UserPanel.UserList.RIO_4ListItem);
            RecreateObjects(WaC, lvl);
            this.UIMap.ARM_AdminWindow.SetFocus();
            Keyboard.SendKeys("{F4}", ModifierKeys.Alt);
        }

        public void DeleteUser(int WaC, int lvl, WpfListItem listitem)
        {
            InputLog("Проверим существование пользователя", lvl); 
            if (listitem.Exists)
            {
                InputLog("Выберем его", lvl + 1);
                Mouse.Click(listitem);
                InputLog("Нажмём кнопку удалить", lvl + 1);
                Mouse.Click(this.UIMap.ARM_AdminWindow.MasterWindow.UserPanel.DeleteUserButton);
                InputLog("Дождёмся подтверждения", lvl + 1);
                this.UIMap.AcceptWindow.WaitForControlExist(10 * WaC);
                InputLog("Подтвердим удаление", lvl + 1);
                Mouse.Click(this.UIMap.AcceptWindow.Acc_YesWindow.YesButton);
                InputLog("Ждём подтверждения удаления", lvl + 1);
                this.UIMap.InfoWindow.WaitForControlExist(100 * WaC);
                InputLog("Нажмём ОК", lvl + 1);
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
            InputLog("Раскроем панель пользователей", lvl);
            Mouse.Click(this.UIMap.ARM_AdminWindow.MasterWindow.NavigationPanel.UsersButton);
            InputLog("Дождёмся появления кнопки Добавить пользователя", lvl);
            this.UIMap.ARM_AdminWindow.AddUserButton.WaitForControlExist(10 * WaC);
            InputLog("Нажмём кнопку Добавить пользователя", lvl);
            Mouse.Click(this.UIMap.ARM_AdminWindow.AddUserButton);
            InputLog("Дождёмся отстройки формы создания пользователя", lvl);
            this.UIMap.ARM_AdminWindow.MasterWindow.EditUserPanel.SelectPersonButton.WaitForControlExist(10 * WaC);
            InputLog("Нажмём кнопку Выбор сотрудника", lvl);
            Mouse.Click(this.UIMap.ARM_AdminWindow.MasterWindow.EditUserPanel.SelectPersonButton);
            InputLog("Дождёмся появления формы создания сотрудника", lvl);
            this.UIMap.EmployeeWindow.WaitForControlExist(10 * WaC);
            InputLog("Проверим существует ли сотрудник "+ UserName, lvl);
            if (!listitem.Exists)
            {
                InputLog("Раскроем поле ввод ФИО", lvl);
                Mouse.Click(this.UIMap.EmployeeWindow.AddEmployeeButton);
                InputLog("Ввёдем ФИО", lvl);
                this.UIMap.EmployeeWindow.FIOEmployeeEdit.Text = UserName;
                InputLog("Добавим сотрудника", lvl);
                Mouse.Click(this.UIMap.EmployeeWindow.AddUserButton);
            }
            InputLog("Выберем сотрудника", lvl);
            Mouse.Click(listitem);
            InputLog("Подтвердим выбор сотрудника", lvl);
            Mouse.Click(this.UIMap.EmployeeWindow.SelectButton);
            InputLog("Введем логин", lvl);
            this.UIMap.ARM_AdminWindow.MasterWindow.EditUserPanel.LoginEdit.Text = UserName;
            InputLog("Введем пароль", lvl);
            this.UIMap.ARM_AdminWindow.MasterWindow.EditUserPanel.PasswordEdit.Text = UserName;
            InputLog("Введем повторно пароль", lvl);
            this.UIMap.ARM_AdminWindow.MasterWindow.EditUserPanel.RePasswordEdit.Text = UserName;
            InputLog("Проверим RadioButton", lvl);
            if (!this.UIMap.ARM_AdminWindow.MasterWindow.EditUserPanel.AsiAutoStartRadioButton.Selected)
            {
                InputLog("Установим значение на АС Инспектора", lvl);
                this.UIMap.ARM_AdminWindow.MasterWindow.EditUserPanel.AsiAutoStartRadioButton.Selected = true;
            }
            InputLog("Сохраним пользователя", lvl);
            Mouse.Click(this.UIMap.ARM_AdminWindow.MasterWindow.EditUserPanel.SaveUserButton);   
            if (Try == 0)
            {
                InputLog("Дождёмся окна ввода данных АИБа", lvl);
                this.UIMap.ISA_Window.WaitForControlExist(30 * WaC);
                InputLog("Ввёдем логин", lvl);
                this.UIMap.ISA_Window.IsaLoginEdit.Text = SchemaName + "_ISA";
                InputLog("Ввёдем пароль", lvl);
                this.UIMap.ISA_Window.IsaPasswordEdit.Text = SchemaName + "_ISA";
                InputLog("Подтвердим ввёденные данные", lvl);
                GetScreen("BeforePressOK");
                Mouse.Click(this.UIMap.ISA_Window.OKButton);
            }
            InputLog("Дождёмся создания пользователя", lvl);
            this.UIMap.UserCreatedWindow.OKWindow2.OKButton.WaitForControlExist(30 * WaC);
            InputLog("Подтвердим выполнение", lvl);
            Mouse.Click(this.UIMap.UserCreatedWindow.OKWindow2.OKButton);
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
            Keyboard.SendKeys("{DOWN}");
            InputLog("Раскроем второй уровень дерева элементов", lvl);
            Keyboard.SendKeys("{RIGHT}");
            InputLog("Дождёмся его появления", lvl);
            this.UIMap.ARM_AdminWindow.Tree.TreeLvL1.TreeLvL2.WaitForControlExist(1 * WaC);
            InputLog("В зависимости от выбранного элемента выберем 117 или 118 ID ", lvl);
            if (this.UIMap.ARM_AdminWindow.MasterWindow.TuPropWindow.RegionsComboBox.SelectedIndex == 118 ||
                    this.UIMap.ARM_AdminWindow.MasterWindow.TuPropWindow.RegionsComboBox.SelectedIndex == -1)
            {
                InputLog("Спустимся ниже", lvl);
                Keyboard.SendKeys("{DOWN}");
            }
            else
            {
                InputLog("Спустимся ниже", lvl);
                Keyboard.SendKeys("{DOWN}");
                InputLog("Спустимся ниже", lvl);
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
            if (//Environment.OSVersion.ToString() != "Microsoft Windows NT 6.1.7601 Service Pack 1" // Пока такое условие, нужно придумать способо обхода
                this.TestContext.Properties["AgentName"].ToString() == "ASI-TST-12-2"
                )
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
                Thread.Sleep(10000);
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
            Mouse.Click(this.UIMap.OKWindow.OKButton);
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
                InputLog(Temp + DirectoryName + " \r\n " +
                    LocalAppData + "\r\n" + InputLanguage.CurrentInputLanguage.ToString()
                    + "\r\n" + InputLanguage.DefaultInputLanguage.ToString()
                    + CultureInfo.CurrentCulture.ToString() + "\r\n"
                    + CultureInfo.CurrentUICulture.ToString() + "\r\n", 0);
                InputLog("Начнём подготовку к экспресс-тестированию", 0);

                //PrepareAsi(WaC, 1, "ASISTA_UI");

               // DownloadSpr(WaC, 1);

                CalcSpr(WaC, 1);
            }
            catch (Exception e)
            {
                InputLog(e.Message, 0);
                GetScreen("Error");
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
