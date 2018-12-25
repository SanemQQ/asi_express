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
using System.Runtime.InteropServices;


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
        readonly string InnBorrower = "123456700063";
        readonly string NameBorrower = "118_New";
        readonly string KppBorrower = "123456789";
        readonly string OgrnBorrower = "313132804400022";
        ushort IdCurrentLanguage ;
        readonly string lang_str = "00000409";
        int ret;
        readonly string Mode = "Debug";


        #region SwapLang
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern bool PostMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        static extern int LoadKeyboardLayout(string pwszKLID, uint Flags);

        private enum KeyboardLayoutFlags : uint
        {
            KLF_ACTIVATE = 0x00000001,
            KLF_SETFORPROCESS = 0x00000100
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint ActivateKeyboardLayout(uint hkl, KeyboardLayoutFlags Flags);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        private static extern ushort GetKeyboardLayout([In] int idThread);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowThreadProcessId(
            [In] IntPtr hWnd,
            [Out] [Optional] IntPtr lpdwProcessId);

        private ushort GetKeyboardLayout()
        {
            return GetKeyboardLayout(GetWindowThreadProcessId(GetForegroundWindow(), IntPtr.Zero));
        }

        #endregion SwapLang

        [TestInitialize]
        public void TestStartup()
        {
            if (!Directory.Exists(Temp + DirectoryName))
            {
                Directory.CreateDirectory(Temp + DirectoryName);
            }
            if (!File.Exists(Temp + DirectoryName + @"\express.log"))
            {
                File.AppendAllText(Temp + DirectoryName + @"\express.log", Temp + DirectoryName + " \r\n " +
                                                                            LocalAppData + "\r\n" + InputLanguage.CurrentInputLanguage.ToString()
                                                                            + "\r\n" + InputLanguage.DefaultInputLanguage.ToString()
                                                                            + CultureInfo.CurrentCulture.ToString()
                                                                            + CultureInfo.CurrentUICulture.ToString());
            }
            else
            {
                File.AppendAllText(Temp + DirectoryName + @"\express_" + DateTime.UtcNow.ToShortDateString().ToString(lang) + ".log", "");
                File.Copy(Temp + DirectoryName + @"\express.log",
                            Temp + DirectoryName + @"\express_" + DateTime.Now.ToString("ddMMyyyy_HHmmss") + ".log");
                File.AppendAllText(Temp + DirectoryName + logfile, Temp + DirectoryName + " \r\n " +
                                                                            LocalAppData + "\r\n" + InputLanguage.CurrentInputLanguage.ToString()
                                                                            + "\r\n" + InputLanguage.DefaultInputLanguage.ToString()
                                                                            + CultureInfo.CurrentCulture.ToString()
                                                                            + CultureInfo.CurrentUICulture.ToString());
            }
            // Справочник точек, по которым будет проводиться взаимодействие с некоторыми объектами
            //RibbonUpperButtons
            Dots.Add("InfoManag", new Point(305, 45));
            Dots.Add("ProccesInfoStat", new Point(670, 45));
            Dots.Add("ProccesInfoMob", new Point(470, 45));
            Dots.Add("PrepareResultMob", new Point(640, 45));


            //RibbonButtons
            Dots.Add("AFSBButton", new Point(165, 80));
            Dots.Add("SPRButton", new Point(55, 80));
            Dots.Add("InfoBorrower", new Point(745, 80));
            Dots.Add("AddBorrowerFromReports", new Point(60, 80));
            Dots.Add("AddCustomBorrower", new Point(125, 80));
            Dots.Add("CreateIndRep", new Point(270, 80));
            Dots.Add("CreateAkt", new Point(535, 80));

            //LeftPanel
            Dots.Add("Material", new Point(20, 335));
            Dots.Add("PinMaterial", new Point(305, 160));
            Dots.Add("FirstFolderMaterial", new Point(150, 215));

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
            Dots.Add("UpperJanuary", new Point(110, 250));
            Dots.Add("Bottom2015", new Point(195, 410));

            //CalcSPR
            Dots.Add("SprAFSBCalc", new Point(65, 250));

            //FormSetQuest
            Dots.Add("AddTask", new Point(95, 205));
            Dots.Add("ChangeCode", new Point(445, 230));
            Dots.Add("ChangeEmp", new Point(1200, 230));
            Dots.Add("131Code", new Point(683, 478));
            Dots.Add("FirstEmp", new Point(595, 365));
            Dots.Add("SecondEmp", new Point(595, 380));
            Dots.Add("SaveCode", new Point(50, 80));

            //DistribTask
            Dots.Add("SaveDisk", new Point(380, 210));
            Dots.Add("OpenFolder", new Point(302, 232));
            Dots.Add("FirstTask", new Point(95, 225));
            Dots.Add("FirstFolder", new Point(350, 250));

            //AddCustomBorrower
            Dots.Add("EditName", new Point(340, 265));
            Dots.Add("TypeBorrower", new Point(340, 300));
            Dots.Add("TypeBorrowerYl", new Point(340, 325));
            Dots.Add("TypeConnect", new Point(340, 330));
            Dots.Add("TypeConnectSsyda", new Point(340, 355));
            Dots.Add("DropFocus", new Point(235, 235)); //123456700063
            Dots.Add("FullName", new Point(420, 450));
            Dots.Add("INN", new Point(420, 480));
            Dots.Add("KPP", new Point(420, 510));
            Dots.Add("OGRN", new Point(420, 540));
            Dots.Add("IP", new Point(510, 540));
            Dots.Add("SelectAdress", new Point(1800, 600));
            Dots.Add("EqualAdress", new Point(210, 625));
            Dots.Add("AddAccounts", new Point(205, 255));
            

            //AddAdress
            Dots.Add("Paremeter_0", new Point(940, 460)); // Регион
            Dots.Add("Paremeter_1", new Point(940, 485)); // Район
            Dots.Add("Paremeter_2", new Point(940, 510)); //
            Dots.Add("Paremeter_3", new Point(940, 535)); //
            Dots.Add("Paremeter_4", new Point(940, 560)); //
            Dots.Add("Paremeter_5", new Point(940, 585)); //
            Dots.Add("Paremeter_6", new Point(940, 610)); //

            //SetTextOnCreateMaterial
            Dots.Add("SetText", new Point(1020, 440));
            //AddAccountsBorrower
            Dots.Add("AcceptAccountsBorrower", new Point(905, 840));

            //FormIndReport
            Dots.Add("IntroPart", new Point(140, 290)); 
            Dots.Add("ChangeIntroPart", new Point(220, 345));

            //FormAktReport
            Dots.Add("EndingPart", new Point(160, 440));
            Dots.Add("ChangeEndingPart", new Point(210, 495));
            Dots.Add("OKButton", new Point(960, 345));
        }

        public void InputLog(string txt, int lvl)  // txt - строка, которую передают для логирования, а lvl - это уровень вложенности
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
        } // Сделать скриншот

         [TestCleanup]
        public void MyTestCleanup()
        {
            Shutdown_asi();
            Thread.Sleep(10000);
            Clear_cache();
            ZipFile.CreateFromDirectory(Temp + DirectoryName, Temp + DirectoryName + this.TestContext.Properties["AgentName"].ToString() + DateTime.Now.ToString("ddMMyyyy_HHmm") + this.TestContext.CurrentTestOutcome.ToString() + ".zip");
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
        } // Открытие формы загрузки из АФСБ

        public void DownloadSpr(int WaC, int lvl)
        {
            InputLog("Открываем вкладку загрузки АФСБ", lvl);
            OpenAFSB(WaC, lvl + 1);
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
        } // Загрузка справочной информации

        public void CalcSpr(int WaC, int lvl) // Загрузка справочников из АФСБ
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
            GetScreen("Finish_calc");
            InputLog("Закрываем", lvl);
            Thread.Sleep(5 * WaC);
            Mouse.Click(this.UIMap.LogWindow.CloseWindow.CloseButton);
        }

        public void CreateQuestions(int WaC, int lvl) // Создание вопросов
        {
            InputLog("Дождёмся открытия формы Формирования заданий", lvl);
            this.UIMap.ASI_Window.CreateTaskWindow.WaitForControlExist(60 * WaC);
            InputLog("Добавим задачу", lvl);
            Mouse.Click(Dots["AddTask"]);
            Thread.Sleep(1 * WaC);
            InputLog("Выберем проверяемы код для этой задачи", lvl);
            Mouse.Click(Dots["ChangeCode"]);
            this.UIMap.SelectCode.WaitForControlExist(60 * WaC);
            InputLog("Выбираем 131 код", lvl);
            Mouse.Click(Dots["131Code"]);
            Thread.Sleep(1 * WaC);
            InputLog("Подтвердим выбор", lvl);
            Mouse.Click(this.UIMap.SelectCode.OKWindow.OKButton);
            InputLog("Дождёмся уничтожения окна", lvl);
            this.UIMap.SelectCode.WaitForControlNotExist(60 * WaC);
            InputLog("выберем сотрудников", lvl);
            Mouse.Click(Dots["ChangeEmp"]);
            this.UIMap.FormSprEmp.WaitForControlExist(60 * WaC);
            InputLog("Выберем первого сотрудника", lvl);
            Mouse.Click(Dots["FirstEmp"]);
            Thread.Sleep(1 * WaC);
            InputLog("Выберем второго сотрудника", lvl);
            Mouse.Click(Dots["SecondEmp"]);
            InputLog("Подтвердим выбор", lvl);
            Mouse.Click(this.UIMap.FormSprEmp.OKWindow.OKButton);
            this.UIMap.FormSprEmp.WaitForControlNotExist(60 * WaC);
            InputLog("Сохраним изменения", lvl);
            Mouse.Click(Dots["SaveCode"]);
            InputLog("Подтвердим выбор", lvl);
            this.UIMap.AcceptWindow.Acc_YesWindow.YesButton.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.AcceptWindow.Acc_YesWindow.YesButton);
        }

        public void DistribTask(int WaC, int lvl)
        {
            InputLog("Дожидаемся открытия формы Распределения задач", lvl);
            this.UIMap.ASI_Window.GivingTaskWindow.WaitForControlExist(60 * WaC);
            InputLog("Подтверждаем создание структуры акта по умолчанию", lvl);
            this.UIMap.AcceptWindow.Acc_YesWindow.YesButton.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.AcceptWindow.Acc_YesWindow.YesButton);
            Thread.Sleep(5 * WaC);
            InputLog("Раскрываем папку вводная часть", lvl);
            Mouse.DoubleClick(Dots["OpenFolder"]);
            Thread.Sleep(1 * WaC);
            InputLog("Перетащим вопрос внутрь папки", lvl);
            Mouse.Move(Dots["FirstTask"]);
            Thread.Sleep(1 * WaC);
            Mouse.StartDragging();
            Mouse.StopDragging(Dots["FirstFolder"]);
            InputLog("Перетащили", lvl);
            InputLog("Ждём появления формы настройки задачи", lvl);
            this.UIMap.SettingTask.WaitForControlExist(60 * WaC);
            Keyboard.SendKeys(this.UIMap.SettingTask.ClarifTaskWindow.ClarifTaskMemo, "Hello World");
            GetScreen("HelloWorld_in_settingtask");
            InputLog("Сохраним изменения в задаче", lvl);
            this.UIMap.SettingTask.SaveWindow.SaveButton.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.SettingTask.SaveWindow.SaveButton);
            this.UIMap.SettingTask.WaitForControlNotExist(60 * WaC);
            InputLog("Сохраним структуру", lvl);
            Mouse.Click(Dots["SaveDisk"]);
            Thread.Sleep(20 * WaC);
        } //Распределение задач

        public void DownloadReports(int WaC, int lvl)
        {
            InputLog("Открываем вкладку загрузки АФСБ", lvl);
            OpenAFSB(WaC, lvl + 1);
            InputLog("Выбрали форму 101", lvl);
            Mouse.Click(Dots["101F"]);
            Thread.Sleep(1 * WaC);
            InputLog("Выбрали форму 102", lvl);
            Mouse.Click(Dots["102F"]);
            Thread.Sleep(1 * WaC);
            InputLog("Выбрали форму 117", lvl);
            Mouse.Click(Dots["117F"]);
            Thread.Sleep(1 * WaC);
            InputLog("Выбрали форму 118", lvl);
            Mouse.Click(Dots["118F"]);
            Thread.Sleep(1 * WaC);
            InputLog("Выбрали форму 501", lvl);
            Mouse.Click(Dots["501F"]);
            Thread.Sleep(1 * WaC);
            InputLog("Открываем нижний календарь", lvl);
            Mouse.Click(Dots["BottomCalendar"]);
            Thread.Sleep(1 * WaC);
            InputLog("Выбираем 2015 год", lvl);
            Mouse.Click(Dots["Bottom2015"]);
            Thread.Sleep(1 * WaC);
            InputLog("Закрываем нижний календарь", lvl);
            Mouse.Click(Dots["BottomCalendar"]);
            Thread.Sleep(1 * WaC);

            InputLog("Открываем верхний календарь", lvl);
            Mouse.Click(Dots["UpperCalendar"]);
            Thread.Sleep(1 * WaC);
            InputLog("Выбираем Январь", lvl);
            Mouse.Click(Dots["UpperJanuary"]);
            Thread.Sleep(1 * WaC);
            InputLog("Закрываем верхний календарь", lvl);
            Mouse.Click(Dots["UpperCalendar"]);
            Thread.Sleep(1 * WaC);

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
        } // Загрузка форм отчётности

        public void AddFile(int WaC, int lvl)
        {
            InputLog("Раскрываем дерево материалов", lvl);
            Mouse.Click(Dots["Material"]);
            Thread.Sleep(6 * WaC);
            InputLog("Зафиксируем его, чтобы не убежал", lvl);
            Mouse.Click(Dots["PinMaterial"]);
            InputLog("Раскроем контекстное меню", lvl);
            Mouse.Click(MouseButtons.Right, ModifierKeys.None, Dots["FirstFolderMaterial"]);
            if (this.UIMap.ContextMenuMaterial.Exists)
            {
                InputLog("Ткнём в элемент загрузки файла из внешней среды", lvl);
                Mouse.Click(this.UIMap.ContextMenuMaterial.MenuItem.UploadFileMenuItem);
            }
            InputLog("Дождёмся диалога открытия файла", lvl);
            this.UIMap.OpenFileDialog.WaitForControlExist(60 * WaC);
            InputLog("Ввёдем имя файла", lvl);
            Keyboard.SendKeys(this.UIMap.OpenFileDialog.FileNameWindow.FileNameEdit, this.TestContext.TestDeploymentDir + "\\" + this.TestContext.Properties["File"].ToString() + ".txt");
            InputLog("Подтвердим выбор файла", lvl);
            Mouse.Click(this.UIMap.OpenFileDialog.OpenWindow.OpenButton);
            this.UIMap.OpenFileDialog.WaitForControlNotExist(60 * WaC);
            InputLog("Ввёдем в фильтр имя файла", lvl);
            Thread.Sleep(5 * WaC);
            this.UIMap.ASI_Window.MaterialWindow.MaterialClient.MaterialPanel.MaterialCustomTree.SearchEdit.Text = this.TestContext.Properties["File"].ToString();
            Thread.Sleep(5 * WaC);
            InputLog("Убедимся, что файл, на втором уровне существует", lvl);
            this.UIMap.ASI_Window.MaterialWindow.MaterialClient.MaterialPanel.MaterialCustomTree.Tree.ASITreeFirstlvl.ASITreeSecondlvl.ItemInSecondlvl.WaitForControlExist(60 * WaC);
            InputLog("Очистим фильтр", lvl);
            Mouse.Click(this.UIMap.ASI_Window.MaterialWindow.MaterialClient.MaterialPanel.MaterialCustomTree.ClearFilter);
            Thread.Sleep(5 * WaC);
            InputLog("Сделаем скриншот", lvl);
            GetScreen("AddedFile");
            InputLog("Снимем фиксацию с дерева материалов", lvl);
            Mouse.Click(Dots["PinMaterial"]);
        } // Добавление файла в дерево материалов

        public void AddBorrower(int WaC, int lvl)
        {
            InputLog("Перейдем в раздел \"Обработка информации\"", lvl);
            Mouse.Click(Dots["ProccesInfoStat"]);
            Thread.Sleep(3 * WaC);
            InputLog("Откроем форму с заёмщиками", lvl);
            Mouse.Click(Dots["InfoBorrower"]);
            InputLog("Ждём", lvl);
            this.UIMap.ASI_Window.BorrowerWindow.WaitForControlExist(60 * WaC);
            Thread.Sleep(5 * WaC);
            InputLog("Добавим заёмщиков из отчётности", lvl);
            Mouse.Click(Dots["AddBorrowerFromReports"]);
            InputLog("Дождёмся формы с выбором дат", lvl);
            this.UIMap.SelectDateWindow.WaitForControlExist(60 * WaC);
            InputLog("Подтвердим выбор", lvl);
            Mouse.Click(this.UIMap.SelectDateWindow.OKWindow.OKButton);
            InputLog("Ждём появления формы с заёмщиками из отчётности", lvl);
            this.UIMap.BorrowerListWindow.WaitForControlExist(60 * WaC);
            InputLog("Подтвердим выбор", lvl);
            Mouse.Click(this.UIMap.BorrowerListWindow.OKWindow.OKButton);
            InputLog("Ждём информации по добавлению заёмщиков", lvl);
            this.UIMap.InfoWindow.OKWindow.OKButton.WaitForControlExist(300 * WaC);
            InputLog("Подтвердим успешность операции", lvl);
            Mouse.Click(this.UIMap.InfoWindow.OKWindow.OKButton);
            InputLog("Сделаем снимок", lvl);
            GetScreen("Added_Borrower");
        } // Добавление заёмщиков из отчётности

        public void AddCustomBorrower(int WaC, int lvl)
        {
            InputLog("Перейдем на первый этап создания заёмщика", lvl);
            AddCustomBorrower1Stage(WaC, lvl + 1);
            InputLog("Перейдем на второй этап создания заёмщика", lvl);
            AddCustomBorrower2Stage(WaC, lvl + 1);
        } // Добавление заёмщика вручную

        public void CheckLanguage(int lvl)
        {
            if (Mode == "Debug")
            {
                InputLog("Проверим раскладку перед вводом данных", lvl);
            }
            IdCurrentLanguage = GetKeyboardLayout();
            ret = LoadKeyboardLayout(lang_str, 1);
            if (IdCurrentLanguage != 1033)
            {
                if (Mode == "Debug")
                {
                    InputLog("Придётся сменить раскладку на en-US, старая " + IdCurrentLanguage.ToString(), lvl + 1);
                }
                PostMessage(GetForegroundWindow(), 0x50, 1, ret);
            }
            if (Mode == "Debug")
            {
                InputLog("Вышли из проверки", lvl);
            }
        } // проверка языка

        public void AddTextOnList(int WaC, int lvl,string txt)
        {
            InputLog("Ткнём в центр листа", lvl);
            Mouse.Click(Dots["SetText"]);
            InputLog("Проверим язык", lvl);
            CheckLanguage(lvl + 1);
            InputLog("Ввёдем текст", lvl);
            Keyboard.SendKeys(txt);
            Thread.Sleep(1 * WaC);
            InputLog("Перейдем на следующую страницу", lvl);
            Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.NextWindow.NextButton);

        } // отдельный метод добавления текста на страницу

        public void AddCustomBorrowerAbsData(int WaC, int lvl)
        {
            InputLog("Добавим счета для заёмщика", lvl);
            Mouse.Click(Dots["AddAccounts"]);
            InputLog("Проверим язык", lvl);
            CheckLanguage(lvl + 1);
            InputLog("Дождёмся появления формы добавления счетов", lvl);
            this.UIMap.SelectAccBorrower.WaitForControlExist(60 * WaC);
            this.UIMap.SelectAccBorrower.SetFocus();
            Keyboard.SendKeys("{TAB}{TAB}{TAB}");
            InputLog("Ввёдем имя заёмщика", lvl);
            Keyboard.SendKeys(NameBorrower);
            InputLog("Покажем все счета из данных", lvl);
            Mouse.Click(this.UIMap.SelectAccBorrower.ShowAccountsWindow.ShowAccountButton);
            Thread.Sleep(4 * WaC);
            InputLog("Перенёсем все найденные счета в правую часть", lvl);
            GetScreen("AllAccountFromAbs");
            Mouse.Click(Dots["AcceptAccountsBorrower"]);
            InputLog("Согласимся с перенесенными данными", lvl);
            Mouse.Click(this.UIMap.SelectAccBorrower.OKWindow.OKButton);
            this.UIMap.AcceptWindow.WaitForControlExist(60 * WaC);
            InputLog("Подтвердим действие", lvl);
            Mouse.Click(this.UIMap.AcceptWindow.Acc_YesWindow.YesButton);
            Thread.Sleep(5 * WaC);
            InputLog("Проверим наличие окна", lvl);
            if (this.UIMap.SelectAccBorrower.Exists)
            {
                InputLog("Попросим закрыться ещё раз", lvl +1);
                Mouse.Click(this.UIMap.SelectAccBorrower.OKWindow.OKButton);
            }
            Thread.Sleep(10 * WaC);
            InputLog("Ждём когда форма исчезнет", lvl);
            this.UIMap.SelectAccBorrower.WaitForControlNotExist(60 * WaC);
            InputLog("Перейдем на следующую страницу", lvl);
            Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.NextWindow.NextButton);
        } // метод добавления данных счетов заёмщика при создании заёмщика вручную

        public void AddCustomBorrowerCalcAbsData(int WaC, int lvl, string txt)
        {
            InputLog("Дождёмся появления кнопки рассчитать на основе АБС", lvl);
            this.UIMap.ASI_Window.AddCustomBorrowerWindow.CalcFromAbsWindow.CalcFromAbsButon.WaitForControlExist(60 * WaC);
            InputLog("Произведём расчёт", lvl);
            Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.CalcFromAbsWindow.CalcFromAbsButon);
            CheckLanguage(lvl + 1);
            InputLog("Ждём появления формы с выбором дат", lvl);
            this.UIMap.SelectDateAbsWindow.WaitForControlExist(60 * WaC);
            InputLog("На эту отчётную дату", lvl);
            Mouse.Click(this.UIMap.SelectDateAbsWindow.OnThatDayWindow.OnThatDayRadioButton);
            InputLog("Подтвердим", lvl);
            Mouse.Click(this.UIMap.SelectDateAbsWindow.OKWindow.OKButton);
            this.UIMap.InfoWindow.WaitForControlExist(60 * WaC);
            InputLog("Подтвердим выполнение расчёта", lvl);
            Mouse.Click(this.UIMap.InfoWindow.OKWindow.OKButton);
            InputLog("Сделам снимок", lvl);
            GetScreen("Calc_From_Abs");
            InputLog("Добавим текст", lvl);
            AddTextOnList(WaC, lvl + 1, txt);
        } // метод для расчета на листе добавленных счетов при создании заёмщика вручную

        public void AddCustomBorrowerCreateReport(int WaC, int lvl)
        {
            InputLog("Ждём открытия формы создания материалов заёмщика", lvl);
            this.UIMap.ASI_Window.AddCustomBorrowerWindow.CreateReportWindow.CreateReportButton.WaitForControlExist(60 * WaC);
            InputLog("Создадим", lvl);
            Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.CreateReportWindow.CreateReportButton);
            this.UIMap.AcceptWindow.WaitForControlExist(60 * WaC);
            InputLog("Подтвердим", lvl);
            Mouse.Click(this.UIMap.AcceptWindow.Acc_YesWindow.YesButton);
            Thread.Sleep(2 * WaC);
            this.UIMap.AcceptWindow.WaitForControlExist(60 * WaC);
            InputLog("Подтвердим ещё раз", lvl);
            Mouse.Click(this.UIMap.AcceptWindow.Acc_YesWindow.YesButton);
            this.UIMap.AcceptWindow.WaitForControlNotExist(60 * WaC);
            InputLog("Ждём поялвения Waiter'а", lvl);
            this.UIMap.RefreshSheetBorrower.WaitForControlExist(60 * WaC);
            this.UIMap.RefreshSheetBorrower.WaitForControlNotExist(600 * WaC);
            InputLog("Ждём появления отчёта", lvl);
            this.UIMap.ReportBorrower.WaitForControlExist(400 * WaC);
            InputLog("Заложили время на \"пролагивание отчёта\"", lvl);
            Thread.Sleep(120 * WaC);
            InputLog("Закроем отчёт", lvl);
            Keyboard.SendKeys("{F4}", ModifierKeys.Alt);
            InputLog("Дождёмся его уничтожения", lvl);
            this.UIMap.ReportBorrower.WaitForControlNotExist(120 * WaC);
            InputLog("Перейдем на другую страницу", lvl);
            Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.NextWindow.NextButton);
            InputLog("Завершим создание заёмщика", lvl);
            this.UIMap.ASI_Window.AddCustomBorrowerWindow.ReadyWindow.ReadyButton.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.ReadyWindow.ReadyButton);
        } // метод для создания материалов по заёмщику

        public void AddCustomBorrower2Stage(int WaC,int lvl)
        {
            for (int i = 0; i <=13; i++)
            {
                InputLog("Попали в цикл", lvl);
                switch (i)
                {
                    case 1:
                        InputLog("Выполним добавление счетов", lvl);
                        AddCustomBorrowerAbsData(WaC, lvl + 1);
                        break;
                    case 2:
                        InputLog("Рассчитаем значения по АБС", lvl);
                        AddCustomBorrowerCalcAbsData(WaC, lvl + 1,"Hello World " + i.ToString());
                        break;
                    case 10:
                        InputLog("Просто перелистнём страницу, сделать ничего не сможем =(", lvl);
                        Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.NextWindow.NextButton);
                        break;
                    case 11:
                        InputLog("Просто перелистнём страницу, сделать ничего не сможем =(", lvl);
                        Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.NextWindow.NextButton);
                        break;
                    case 12:
                        InputLog("Просто перелистнём страницу, сделать ничего не сможем =(", lvl);
                        Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.NextWindow.NextButton);
                        break;
                    case 13:
                        InputLog("Попытаемся сформировать материалы по заёмщику", lvl);
                        AddCustomBorrowerCreateReport(WaC, lvl + 1);
                        break;
                    default:
                        InputLog("Добавим текст в " + i.ToString()+" раз", lvl);
                        AddTextOnList(WaC, lvl + 1, "Hello World " + i.ToString());                        
                        break;
                }
                Thread.Sleep(1 * WaC);                
            }

        } // сгруппированный этап для всех листов, кроме первого

        public void AddCustomBorrower1Stage(int WaC, int lvl)
        {
            InputLog("Проверим язык", lvl);
            CheckLanguage(lvl + 1);
            InputLog("Откроем вкладку обработка информации", lvl);
            Mouse.Click(Dots["ProccesInfoMob"]);
            Thread.Sleep(3 * WaC);
            InputLog("Откроем информацию о заёмщиках", lvl);
            Mouse.Click(Dots["InfoBorrower"]);
            this.UIMap.ASI_Window.BorrowerWindow.WaitForControlExist(60 * WaC);
            Thread.Sleep(3 * WaC);
            InputLog("Создадим своего заёмщика", lvl);
            Mouse.Click(Dots["AddCustomBorrower"]);
            this.UIMap.ASI_Window.AddCustomBorrowerWindow.WaitForControlExist(120 * WaC);
            Thread.Sleep(15 * WaC);
            InputLog("Раскроем тип заёмщика", lvl);
            Mouse.Click(Dots["TypeBorrower"]);
            Thread.Sleep(5 * WaC);
            InputLog("Выберем ЮЛ", lvl);
            Mouse.Click(Dots["TypeBorrowerYl"]);
            Mouse.Click(Dots["DropFocus"]);
            Thread.Sleep(5 * WaC);
            InputLog("Введём имя", lvl);
            Mouse.Click(Dots["EditName"]);
            Thread.Sleep(5 * WaC);
            Keyboard.SendKeys(NameBorrower);
            Mouse.Click(Dots["DropFocus"]);
            Thread.Sleep(3 * WaC);
            InputLog("Введём тип связи", lvl);
            Mouse.Click(Dots["TypeConnect"]);
            Thread.Sleep(5 * WaC);
            InputLog("Выберем ссуды", lvl);
            Mouse.Click(Dots["TypeConnectSsyda"]);
            Mouse.Click(Dots["DropFocus"]);
            Thread.Sleep(5 * WaC);
            InputLog("Введём полное имя", lvl);
            Mouse.Click(Dots["FullName"]);
            Thread.Sleep(2 * WaC);
            Keyboard.SendKeys(NameBorrower + "_Full");
            Mouse.Click(Dots["IP"]);
            InputLog("Пускай это будет ИП", lvl);
            Thread.Sleep(6 * WaC);
            InputLog("Физ. адрес совпадает с Юр. адресом", lvl);
            Mouse.Click(Dots["EqualAdress"]);
            Thread.Sleep(2 * WaC);
            InputLog("Введём ИНН", lvl);
            Mouse.Click(Dots["INN"]);
            Thread.Sleep(2 * WaC);
            Keyboard.SendKeys(InnBorrower);
            InputLog("Введём КПП", lvl);
            Mouse.Click(Dots["KPP"]);
            Thread.Sleep(2 * WaC);
            Keyboard.SendKeys(KppBorrower);
            InputLog("Введём ОГРН", lvl);
            Mouse.Click(Dots["OGRN"]);
            Thread.Sleep(2 * WaC);
            Keyboard.SendKeys(OgrnBorrower);
            InputLog("Введём адрес", lvl);
            Mouse.Click(Dots["SelectAdress"]);
            this.UIMap.AddAdressWindow.WaitForControlExist(60 * WaC);

            for (int i = 0; i < 7; i++) //i - 7, количество комбиков/эдиток на форме выбора адреса, так как с ними нельзя работать, приходится извращаться
            {
                if(i<3)
                {
                    InputLog("Введем данные адреса", lvl+1);
                    Mouse.Click(Dots["Paremeter_" + i.ToString()]);
                    Keyboard.SendKeys("{DOWN}");
                    Thread.Sleep(1 * WaC);
                }
                else
                {
                    InputLog("Введем данные адреса", lvl + 1);
                    Mouse.Click(Dots["Paremeter_" + i.ToString()]);
                    Keyboard.SendKeys(i.ToString());
                    Thread.Sleep(1 * WaC);
                }
            }
            InputLog("Подтвердим введеные данные", lvl);
            Mouse.Click(this.UIMap.AddAdressWindow.OKWindow.OKButton);
            InputLog("Перейдем на следующую страницу", lvl);
            Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.NextWindow.NextButton);
        } // Первый этап создания заёмщика

        public void OpenReport(int WaC,int lvl,string Report)
        {
            InputLog("Раскроем вкладку подготовка результатов", lvl );
            Mouse.Click(Dots["PrepareResultMob"]);
            Thread.Sleep(1 * WaC);
            InputLog("Отредактируем "+Report, lvl);
            Mouse.Click(Dots[Report]);
            Thread.Sleep(5 * WaC);
        } // Открытие на влкадке информационное обеспечение индвидуального отчёта или акта

        public void CreateIndRep(int WaC, int lvl)
        {
            InputLog("Создадим индивидуальный отчёт", lvl );
            OpenReport(WaC, lvl + 1, "CreateIndRep");
            InputLog("Проверим появление окна с предупреждением о требовании с перезапуском АСИ", lvl );
            if (this.UIMap.WarningWindow.Exists)
            {
                InputLog("Отказываемся от формирования", lvl + 1);
                this.UIMap.WarningWindow.WarnNo_Window.NoButton.WaitForControlExist(60 * WaC);
                Mouse.Click(this.UIMap.WarningWindow.WarnNo_Window.NoButton);
                InputLog("Перезайдем в АСИ", lvl + 1);
                ConnectToWorkingARM(WaC, lvl + 1, "ASIMOB_UI", "RIO");
                InputLog("Создадим индивидуальный отчёт", lvl + 1);
                OpenReport(WaC, lvl + 1, "CreateIndRep");
            }
            this.UIMap.AttentionWindow.WaitForControlExist(60 * WaC);
            InputLog("Игнорируем предупреждение о незагруженности данных", lvl);
            Mouse.Click(this.UIMap.AttentionWindow.Att_NoWindow.NoButton);
            this.UIMap.NameIndRepWindow.WaitForControlExist(60 * WaC);
            InputLog("Введем имя нашего отчёта", lvl);
            this.UIMap.NameIndRepWindow.NameIndRepEdit.Text = "Индивидуальный отчет (РИО)_" + DateTime.Now.ToString("ddMMyyyy_HHmmss");
            Mouse.Click(this.UIMap.NameIndRepWindow.OKButton);
            InputLog("Ждём его построения", lvl);
            this.UIMap.OpenEditor.WaitForControlExist(120 * WaC);
            this.UIMap.OpenEditor.WaitForControlNotExist(120 * WaC);
            this.UIMap.ASI_Window.IndReportWindow.WaitForControlExist(120 * WaC);
            this.UIMap.CreatingWindow.WaitForControlExist(60 * WaC);
            this.UIMap.CreatingWindow.ProgressBar.WaitForControlNotExist(900 * WaC);
            InputLog("Отчёт должен был сформироваться", lvl);
            Thread.Sleep(5 * WaC);
            if (this.UIMap.CreatingWindow.OKButton.Exists)
            {
                InputLog("Метод 1 для нажатия на кнопку ОК", lvl + 1);
                Mouse.Click(this.UIMap.CreatingWindow.OKButton);
            }
            else
            {
                InputLog("Метод 2 для нажатия на кнопку ОК", lvl + 1);
                Mouse.Click(Dots["OKButton"]);
            }
            InputLog("Раскроем первый раздел и воткнём туда какой-нибудь текст", lvl);
            Mouse.Click(MouseButtons.Right, ModifierKeys.None, Dots["IntroPart"]);
            Thread.Sleep(1 * WaC);
            Mouse.Click(Dots["ChangeIntroPart"]);
            this.UIMap.DocumentWindow.WaitForControlExist(60 * WaC);
            InputLog("Ткнём в документ", lvl);
            Mouse.Click(Dots["SetText"]);
            InputLog("Проверим язык", lvl);
            CheckLanguage(lvl + 1);
            InputLog("Вводим текущее время", lvl);
            Keyboard.SendKeys("TODAY_" + DateTime.Now.ToString("ddMMyyyy_HHmmss"));
            InputLog("Сохраним", lvl);
            Mouse.Click(this.UIMap.DocumentWindow.SaveButton);
            this.UIMap.AcceptWindow.Acc_YesWindow.YesButton.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.AcceptWindow.Acc_YesWindow.YesButton);
            InputLog("Да, нужно перестроить зависимые объекты", lvl);
            Thread.Sleep(8 * WaC);
            InputLog("Сделаем скриншот", lvl);
            GetScreen("Create_Changes");
            this.UIMap.DocumentWindow.SetFocus(); // Постараемся установить фокус на окне конструктора
            InputLog("Закроем окно", lvl);
            Keyboard.SendKeys("{F4}", ModifierKeys.Alt);
            this.UIMap.DocumentWindow.WaitForControlNotExist(60 * WaC);
        } // метод для создания индивидуального отчёта

        public void CreateAkt(int WaC, int lvl)
        {
            InputLog("Создадим акт", lvl);
            OpenReport(WaC, lvl + 1, "CreateAkt");
            this.UIMap.EndedCheckWindow.WaitForControlExist(60 * WaC);
            InputLog("Оказывается проверка завершена", lvl);
            Mouse.Click(this.UIMap.EndedCheckWindow.Panel.CreateAktRadioButton);
            this.UIMap.EndedCheckWindow.OKButton.WaitForControlEnabled(60 * WaC);
            InputLog("Да, мы хотим создать акт", lvl);
            Mouse.Click(this.UIMap.EndedCheckWindow.OKButton);
            Thread.Sleep(8 * WaC);
            InputLog("Проверим появление окна с предупреждением о требовании с перезапуском АСИ", lvl);
            if (this.UIMap.WarningWindow.Exists)
            {
                InputLog("Отказываемся от формирования", lvl + 1);
                this.UIMap.WarningWindow.WarnNo_Window.NoButton.WaitForControlExist(60 * WaC);
                Mouse.Click(this.UIMap.WarningWindow.WarnNo_Window.NoButton);
                InputLog("Перезайдем в АСИ", lvl + 1);
                ConnectToWorkingARM(WaC, lvl + 1, "ASIMOB_UI", "RIO");
                InputLog("Создадим акт", lvl + 1);
                OpenReport(WaC, lvl + 1, "CreateAkt");
            }
            this.UIMap.AttentionWindow.WaitForControlExist(60 * WaC);
            InputLog("Игнорируем предупреждение о незагруженности данных", lvl);
            Mouse.Click(this.UIMap.AttentionWindow.Att_NoWindow.NoButton);
            this.UIMap.NameAktWindow.WaitForControlExist(60 * WaC);
            InputLog("Введем имя для акта", lvl);
            this.UIMap.NameAktWindow.NameAktEdit.Text = "Акт проверки (РИО)_" + DateTime.Now.ToString("ddMMyyyy_HHmmss");
            Mouse.Click(this.UIMap.NameAktWindow.OKButton);
            InputLog("Ждём его построения", lvl);
            this.UIMap.OpenEditor.WaitForControlExist(120 * WaC);
            this.UIMap.OpenEditor.WaitForControlNotExist(120 * WaC);
            this.UIMap.ASI_Window.AktReportWindow.WaitForControlExist(120 * WaC);
            this.UIMap.CreatingWindow.WaitForControlExist(60 * WaC);
            this.UIMap.CreatingWindow.ProgressBar.WaitForControlNotExist(900 * WaC);
            InputLog("Отчёт должен был сформироваться", lvl);
            Thread.Sleep(5 * WaC);
            if (this.UIMap.CreatingWindow.OKButton.Exists)
            {
                InputLog("Метод 1 для нажатия на кнопку ОК", lvl + 1);
                Mouse.Click(this.UIMap.CreatingWindow.OKButton);
            }
            else
            {
                InputLog("Метод 2 для нажатия на кнопку ОК", lvl + 1);
                Mouse.Click(Dots["OKButton"]);
            }
            InputLog("Раскроем последний раздел и воткнём туда какой-нибудь текст", lvl);
            Mouse.Click(MouseButtons.Right, ModifierKeys.None, Dots["EndingPart"]);
            Thread.Sleep(1 * WaC);
            Mouse.Click(Dots["ChangeEndingPart"]);
            this.UIMap.DocumentWindow.WaitForControlExist(60 * WaC);
            InputLog("Ткнём в документ", lvl);
            Mouse.Click(Dots["SetText"]);
            InputLog("Проверим язык", lvl);
            CheckLanguage(lvl + 1);
            InputLog("Вводим текущее время", lvl);
             Keyboard.SendKeys("TODAY_" + DateTime.Now.ToString("ddMMyyyy_HHmmss"));
            InputLog("Сохраним", lvl);
            Mouse.Click(this.UIMap.DocumentWindow.SaveButton);
            Thread.Sleep(8 * WaC);
            InputLog("Сделаем скриншот", lvl);
            GetScreen("Create_Changes");
            this.UIMap.DocumentWindow.SetFocus(); // Постараемся установить фокус на окне конструктора
            InputLog("Закроем окно", lvl);
            Keyboard.SendKeys("{F4}", ModifierKeys.Alt);
            this.UIMap.DocumentWindow.WaitForControlNotExist(60 * WaC);
        } // метод для создания акта

        public void Shutdown_asi()
        {
            Process[] p1;
            if (this.TestContext.Properties["ControllerName"].ToString().Contains("localhost"))
            {
                 p1 = Process.GetProcessesByName("Studio");
            }
            else
            {
                p1 = Process.GetProcessesByName("P5");
            }
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
        } // метод выключения АСИ

        public void Clear_cache()
        {
            if (Directory.Exists(LocalAppData + @"\JSC Prognoz"))
            {
                Directory.Delete(LocalAppData + @"\JSC Prognoz", true);
            }
        } // метод для чистки кэша

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

        } // метод запуска прогноза на различных машинах

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
        } // метод запуска АСИ
        
        public void StatsArm(int WaC, int lvl, string SchemaName)
        {
            InputLog("Ждём открытия АРМ Админа на стационарном уровне", lvl);
            this.UIMap.ARM_AdminWindow.WaitForControlExist(60 * WaC);            
            InputLog("Проверим версию системы, если она равна Microsoft Windows NT 6.1.7601 Service Pack 1 пропустим пункт выбора ТУ", lvl);
            if (Environment.OSVersion.ToString() != "Microsoft Windows NT 6.1.7601 Service Pack 1")
            {
                InputLog("Выберем ТУ", lvl);
                SelectTU(WaC, lvl + 1);
            }
            InputLog("Создадим/Пересоздадим соединение с САФСБ", lvl);
            SetAFSB(WaC, lvl + 1);
            InputLog("Создадим пользователя с именем RIO_3", lvl);
            CreateUser(WaC, lvl + 1, this.UIMap.EmployeeWindow.EmployeeList.RIO_3ListItem, "RIO_3", 0, SchemaName);
            InputLog("Удалим пользователя", lvl);
            DeleteUser(WaC, lvl + 1, this.UIMap.ARM_AdminWindow.MasterWindow.UserPanel.UserList.RIO_3ListItem);
            InputLog("Создадим пользователя с именем RIO_4", lvl);
            CreateUser(WaC, lvl + 1, this.UIMap.EmployeeWindow.EmployeeList.RIO_4ListItem, "RIO_4", 1, SchemaName);
            InputLog("Удалим пользователя", lvl);
            DeleteUser(WaC, lvl + 1, this.UIMap.ARM_AdminWindow.MasterWindow.UserPanel.UserList.RIO_4ListItem);
            InputLog("Пересоздадим объекты", lvl);
            RecreateObjects(WaC, lvl);
            InputLog("Закроем АРМ Админа", lvl);
            this.UIMap.ARM_AdminWindow.SetFocus();
            Keyboard.SendKeys("{F4}", ModifierKeys.Alt);
        } // работа внутри АРМа на стационарном уровне

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
                Mouse.Click(this.UIMap.InfoWindow.OKWindow.OKButton);
            }
        } // метод удаления пользователя из списка

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
        } // метод для пересоздания объектов АСИ

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
        } // метод для выбора ТУ

        private void WaiterForMSSQL(int Time)
        {
            if (this.TestContext.Properties["AgentName"].ToString() != "ASI-TST-12-2") { Thread.Sleep(Time*1000); }
        } // специальная ожидалка для мсскл

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
        } // настройка параметров для создания соединения с САФСБ

        public void MobArm(int WaC, int lvl, string SchemaName)
        {
            InputLog("Ждём открытия АРМ Админа на мобильном уровне", lvl);
            this.UIMap.ARM_AdminWindow.WaitForControlExist(30 * WaC);
        } // работа внутри АРМа на мобильном уровне

        public void PrepareAsi(int WaC, int lvl, string SchemaName)
        {
            InputLog("Почистим кэш", lvl);
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
        } // метод для подготовки АСИ к экспресс-тестированию

        public void ConnectToWorkingARM (int WaC, int lvl, string SchemaName,string UserName)
        {
            InputLog("Вырубим АС Инспектора", lvl);
            Shutdown_asi();
            Thread.Sleep(2 * WaC);
            InputLog("Почистим кэш", lvl);
            Clear_cache();
            InputLog("Запустим АСИ ", lvl);
            StartASI(WaC, lvl + 1, SchemaName, UserName);
            InputLog("Дождёмся появления окна АРМа", lvl);
            this.UIMap.ASI_Window.WaitForControlExist(240 * WaC);
            InputLog("Дадим отстроится элементам", lvl);
            Thread.Sleep(10 * WaC);
        } // метод для переподключения к АРМу пользователя/администратора

        #endregion

        [TestMethod]
        [TestProperty("AgentName", "ASI-TST-MS12")]
        [TestProperty("File", "FileToDeploy")]
        [TestProperty("SysLogin", "sa")]
        [TestProperty("SysPassword", "Qwerty1")]
        public void Asi_Express_MSSQL() => Asi_Express_All(1000);

        [TestMethod]
        [TestProperty("AgentName", "ASI-TST-12-2")]
        [TestProperty("File", "FileToDeploy")]
        [TestProperty("SysLogin", "SYSTEM")]
        [TestProperty("SysPassword", "ASITST11")]
        public void Asi_Express_ORACLE() => Asi_Express_All(4000);



        public void Asi_Express_All(int WaC) // WaC - Waiter Coefficient - коэффициент ожидания, который будет корректировать время ожидания между ораклом и мсскл
        {

            try
            {
                InputLog("Начнём подготовку к экспресс-тестированию", 0);


                PrepareAsi(WaC, 1, "ASISTA_UI");
                ConnectToWorkingARM(WaC, 1, "ASISTA_UI", "RIO");
                DownloadSpr(WaC, 1);
                CalcSpr(WaC, 1);
                Shutdown_asi();
                CreateQuestions(WaC, 1);
                DistribTask(WaC, 1);

                DownloadReports(WaC, 1);
                AddFile(WaC, 1);
                AddBorrower(WaC, 1);

                AddCustomBorrower(WaC, 1);
                CheckLanguage(0);
                AddCustomBorrowerCreateReport(WaC, 1);
                AddCustomBorrower(WaC, 1);
                CreateIndRep(WaC,1);
                CreateAkt(WaC, 1);
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
