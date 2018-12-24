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
        }

        // [TestCleanup]
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
        }

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
            this.UIMap.ASI_Window.CreateTaskWindow.WaitForControlExist(60 * WaC);
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
            Mouse.Click(Dots["SaveCode"]);
            this.UIMap.AcceptWindow.Acc_YesWindow.YesButton.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.AcceptWindow.Acc_YesWindow.YesButton);
        }

        public void DistribTask(int WaC, int lvl)
        {
            this.UIMap.ASI_Window.GivingTaskWindow.WaitForControlExist(60 * WaC);
            this.UIMap.AcceptWindow.Acc_YesWindow.YesButton.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.AcceptWindow.Acc_YesWindow.YesButton);
            Thread.Sleep(5 * WaC);
            Mouse.DoubleClick(Dots["OpenFolder"]);
            Thread.Sleep(1 * WaC);
            Mouse.Move(Dots["FirstTask"]);
            Thread.Sleep(1 * WaC);
            Mouse.StartDragging();
            Mouse.StopDragging(Dots["FirstFolder"]);
            this.UIMap.SettingTask.WaitForControlExist(60 * WaC);
            Keyboard.SendKeys(this.UIMap.SettingTask.ClarifTaskWindow.ClarifTaskMemo, "Hello World");
            GetScreen("HelloWorld_in_settingtask");
            this.UIMap.SettingTask.SaveWindow.SaveButton.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.SettingTask.SaveWindow.SaveButton);
            this.UIMap.SettingTask.WaitForControlNotExist(60 * WaC);
            Mouse.Click(Dots["SaveDisk"]);
            Thread.Sleep(20 * WaC);
        }

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
        }

        public void AddFile(int WaC, int lvl)
        {
            Mouse.Click(Dots["Material"]);
            Thread.Sleep(6 * WaC);
            Mouse.Click(Dots["PinMaterial"]);
            Mouse.Click(MouseButtons.Right, ModifierKeys.None, Dots["FirstFolderMaterial"]);
            if (this.UIMap.ContextMenuMaterial.Exists)
            {
                Mouse.Click(this.UIMap.ContextMenuMaterial.MenuItem.UploadFileMenuItem);
            }
            this.UIMap.OpenFileDialog.WaitForControlExist(60 * WaC);
            Keyboard.SendKeys(this.UIMap.OpenFileDialog.FileNameWindow.FileNameEdit, this.TestContext.TestDeploymentDir + "\\" + this.TestContext.Properties["File"].ToString() + ".txt");
            Mouse.Click(this.UIMap.OpenFileDialog.OpenWindow.OpenButton);
            this.UIMap.OpenFileDialog.WaitForControlNotExist(60 * WaC);
            this.UIMap.ASI_Window.MaterialWindow.MaterialClient.MaterialPanel.MaterialCustomTree.SearchEdit.Text = this.TestContext.Properties["File"].ToString();
            Thread.Sleep(5 * WaC);
            this.UIMap.ASI_Window.MaterialWindow.MaterialClient.MaterialPanel.MaterialCustomTree.Tree.ASITreeFirstlvl.ASITreeSecondlvl.ItemInSecondlvl.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.ASI_Window.MaterialWindow.MaterialClient.MaterialPanel.MaterialCustomTree.ClearFilter);
            Thread.Sleep(5 * WaC);
            GetScreen("AddedFile");
            Mouse.Click(Dots["PinMaterial"]);

        }

        public void AddBorrower(int WaC, int lvl)
        {
            Mouse.Click(Dots["ProccesInfoStat"]);
            Thread.Sleep(3 * WaC);
            Mouse.Click(Dots["InfoBorrower"]);
            this.UIMap.ASI_Window.BorrowerWindow.WaitForControlExist(60 * WaC);
            Thread.Sleep(5 * WaC);
            Mouse.Click(Dots["AddBorrowerFromReports"]);
            this.UIMap.SelectDateWindow.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.SelectDateWindow.OKWindow.OKButton);
            this.UIMap.BorrowerListWindow.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.BorrowerListWindow.OKWindow.OKButton);
            this.UIMap.InfoWindow.OKWindow.OKButton.WaitForControlExist(300 * WaC);
            Mouse.Click(this.UIMap.InfoWindow.OKWindow.OKButton);
            GetScreen("Added_Borrower");
        }

        public void AddCustomBorrower(int WaC, int lvl)
        {    
            AddCustomBorrower1Stage(WaC, lvl + 1);
            AddCustomBorrower2Stage(WaC, lvl + 1);
        }

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
        }

        public void AddTextOnList(int WaC, int lvl,string txt)
        {
            Mouse.Click(Dots["SetText"]);
            CheckLanguage(lvl + 1);
            Keyboard.SendKeys(txt);
            Thread.Sleep(1 * WaC);
            Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.NextWindow.NextButton);

        }

        public void AddCustomBorrowerAbsData(int WaC, int lvl)
        {
            Mouse.Click(Dots["AddAccounts"]);
            CheckLanguage(lvl + 1);
            this.UIMap.SelectAccBorrower.WaitForControlExist(60 * WaC);
            this.UIMap.SelectAccBorrower.SetFocus();
            Keyboard.SendKeys("{TAB}{TAB}{TAB}");
            Keyboard.SendKeys(NameBorrower);
            Mouse.Click(this.UIMap.SelectAccBorrower.ShowAccountsWindow.ShowAccountButton);
            Thread.Sleep(4 * WaC);
            Mouse.Click(Dots["AcceptAccountsBorrower"]);
            Mouse.Click(this.UIMap.SelectAccBorrower.OKWindow.OKButton);
            this.UIMap.AcceptWindow.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.AcceptWindow.Acc_YesWindow.YesButton);
            Thread.Sleep(5 * WaC);
            if(this.UIMap.SelectAccBorrower.Exists)
            {
                Mouse.Click(this.UIMap.SelectAccBorrower.OKWindow.OKButton);
            }
            Thread.Sleep(10 * WaC);
            this.UIMap.SelectAccBorrower.WaitForControlNotExist(60 * WaC);
            Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.NextWindow.NextButton);
        }

        public void AddCustomBorrowerCalcAbsData(int WaC, int lvl, string txt)
        {
            this.UIMap.ASI_Window.AddCustomBorrowerWindow.CalcFromAbsWindow.CalcFromAbsButon.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.CalcFromAbsWindow.CalcFromAbsButon);
            CheckLanguage(lvl + 1);
            this.UIMap.SelectDateAbsWindow.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.SelectDateAbsWindow.OnThatDayWindow.OnThatDayRadioButton);
            Mouse.Click(this.UIMap.SelectDateAbsWindow.OKWindow.OKButton);
            this.UIMap.InfoWindow.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.InfoWindow.OKWindow.OKButton);
            AddTextOnList(WaC, lvl + 1, txt);
        }

        public void AddCustomBorrowerCreateReport(int WaC, int lvl)
        {
            try
            {
                this.UIMap.ASI_Window.AddCustomBorrowerWindow.CreateReportWindow.CreateReportButton.WaitForControlExist(60 * WaC);
                Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.CreateReportWindow.CreateReportButton);
                this.UIMap.AcceptWindow.WaitForControlExist(60 * WaC);
                Mouse.Click(this.UIMap.AcceptWindow.Acc_YesWindow.YesButton);
                Thread.Sleep(2 * WaC);
                this.UIMap.AcceptWindow.WaitForControlExist(60 * WaC);
                Mouse.Click(this.UIMap.AcceptWindow.Acc_YesWindow.YesButton);
                this.UIMap.AcceptWindow.WaitForControlNotExist(60 * WaC);
                this.UIMap.RefreshSheetBorrower.WaitForControlExist(60 * WaC);
                this.UIMap.RefreshSheetBorrower.WaitForControlNotExist(600 * WaC);
                this.UIMap.ReportBorrower.WaitForControlExist(400 * WaC);
                Thread.Sleep(120 * WaC);
                Keyboard.SendKeys("{F4}", ModifierKeys.Alt);
                this.UIMap.ReportBorrower.WaitForControlNotExist(120 * WaC);
                Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.NextWindow.NextButton);
                this.UIMap.ASI_Window.AddCustomBorrowerWindow.ReadyWindow.ReadyButton.WaitForControlExist(60 * WaC);
                Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.ReadyWindow.ReadyButton);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public void AddCustomBorrower2Stage(int WaC,int lvl)
        {
            for (int i = 0; i <=13; i++)
            {
                switch (i)
                {
                    case 1:
                        AddCustomBorrowerAbsData(WaC, lvl + 1);
                        break;
                    case 2:
                        AddCustomBorrowerCalcAbsData(WaC, lvl + 1,"Hello World " + i.ToString());
                        break;
                    case 10:
                        Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.NextWindow.NextButton);
                        break;
                    case 11:
                        Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.NextWindow.NextButton);
                        break;
                    case 12:
                        Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.NextWindow.NextButton);
                        break;
                    case 13:
                        AddCustomBorrowerCreateReport(WaC, lvl + 1);
                        break;
                    default:
                        AddTextOnList(WaC, lvl + 1, "Hello World " + i.ToString());
                        break;
                }
                Thread.Sleep(1 * WaC);                
            }

        }

        public void AddCustomBorrower1Stage(int WaC, int lvl)
        {
            CheckLanguage(lvl + 1);
            Mouse.Click(Dots["ProccesInfoMob"]);
            Thread.Sleep(3 * WaC);
            Mouse.Click(Dots["InfoBorrower"]);
            this.UIMap.ASI_Window.BorrowerWindow.WaitForControlExist(60 * WaC);
            Thread.Sleep(3 * WaC);
            Mouse.Click(Dots["AddCustomBorrower"]);
            this.UIMap.ASI_Window.AddCustomBorrowerWindow.WaitForControlExist(120 * WaC);
            Thread.Sleep(15 * WaC);
            Mouse.Click(Dots["TypeBorrower"]);
            Thread.Sleep(5 * WaC);
            Mouse.Click(Dots["TypeBorrowerYl"]);
            Mouse.Click(Dots["DropFocus"]);
            Thread.Sleep(5 * WaC);
            Mouse.Click(Dots["EditName"]);
            Thread.Sleep(5 * WaC);
            Keyboard.SendKeys(NameBorrower);
            Mouse.Click(Dots["DropFocus"]);
            Thread.Sleep(3 * WaC);
            Mouse.Click(Dots["TypeConnect"]);
            Thread.Sleep(5 * WaC);
            Mouse.Click(Dots["TypeConnectSsyda"]);
            Mouse.Click(Dots["DropFocus"]);
            Thread.Sleep(5 * WaC);
            Mouse.Click(Dots["FullName"]);
            Thread.Sleep(2 * WaC);
            Keyboard.SendKeys(NameBorrower + "_Full");
            Mouse.Click(Dots["IP"]);
            Thread.Sleep(6 * WaC);
            Mouse.Click(Dots["EqualAdress"]);
            Thread.Sleep(2 * WaC);
            Mouse.Click(Dots["INN"]);
            Thread.Sleep(2 * WaC);
            Keyboard.SendKeys(InnBorrower);
            Mouse.Click(Dots["KPP"]);
            Thread.Sleep(2 * WaC);
            Keyboard.SendKeys(KppBorrower);
            Mouse.Click(Dots["OGRN"]);
            Thread.Sleep(2 * WaC);
            Keyboard.SendKeys(OgrnBorrower);
            Mouse.Click(Dots["SelectAdress"]);
            this.UIMap.AddAdressWindow.WaitForControlExist(60 * WaC);

            for (int i = 0; i < 7; i++) //i - 7, количество комбиков/эдиток на форме выбора адреса, так как с ними нельзя работать, приходится извращаться
            {
                if(i<3)
                {
                    Mouse.Click(Dots["Paremeter_" + i.ToString()]);
                    Keyboard.SendKeys("{DOWN}");
                    Thread.Sleep(1 * WaC);
                }
                else
                {
                    Mouse.Click(Dots["Paremeter_" + i.ToString()]);
                    Keyboard.SendKeys(i.ToString());
                    Thread.Sleep(1 * WaC);
                }
            }
            Mouse.Click(this.UIMap.AddAdressWindow.OKWindow.OKButton);
            Mouse.Click(this.UIMap.ASI_Window.AddCustomBorrowerWindow.NextWindow.NextButton);
        }

        public void OpenReport(int WaC,int lvl,string Report)
        {
            Mouse.Click(Dots["PrepareResultMob"]);
            Thread.Sleep(1 * WaC);
            Mouse.Click(Dots[Report]);
            Thread.Sleep(5 * WaC);
        }

        public void CreateIndRep(int WaC, int lvl)
        {
            OpenReport(WaC, lvl + 1, "CreateIndRep");
            if (this.UIMap.WarningWindow.Exists)
            {
                this.UIMap.WarningWindow.WarnNo_Window.NoButton.WaitForControlExist(60 * WaC);
                Mouse.Click(this.UIMap.WarningWindow.WarnNo_Window.NoButton);
                ConnectToWorkingARM(WaC, lvl + 1, "ASIMOB_UI", "RIO");
                OpenReport(WaC, lvl + 1, "CreateIndRep");
            }
            this.UIMap.AttentionWindow.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.AttentionWindow.Att_NoWindow.NoButton);
            this.UIMap.NameIndRepWindow.WaitForControlExist(60 * WaC);
            this.UIMap.NameIndRepWindow.NameIndRepEdit.Text = "Индивидуальный отчет (РИО)_" + DateTime.Now.ToString("ddMMyyyy_HHmmss");
            Mouse.Click(this.UIMap.NameIndRepWindow.OKButton);
            this.UIMap.OpenEditor.WaitForControlExist(120 * WaC);
            this.UIMap.OpenEditor.WaitForControlNotExist(120 * WaC);
            this.UIMap.ASI_Window.IndReportWindow.WaitForControlExist(120 * WaC);
            this.UIMap.CreatingWindow.WaitForControlExist(60 * WaC);
            this.UIMap.CreatingWindow.ProgressBar.WaitForControlNotExist(900 * WaC);
            Thread.Sleep(5 * WaC);
            if (this.UIMap.CreatingWindow.OKButton.Exists)
            {
                Mouse.Click(this.UIMap.CreatingWindow.OKButton);
            }
            else
            {
                Mouse.Click(Dots["OKButton"]);
            }
            Mouse.Click(MouseButtons.Right, ModifierKeys.None, Dots["IntroPart"]);
            Thread.Sleep(1 * WaC);
            Mouse.Click(Dots["ChangeIntroPart"]);
            this.UIMap.DocumentWindow.WaitForControlExist(60 * WaC);
            Mouse.Click(Dots["SetText"]);            
            CheckLanguage(lvl + 1);
            Keyboard.SendKeys("TODAY_" + DateTime.Now.ToString("ddMMyyyy_HHmmss"));
            Mouse.Click(this.UIMap.DocumentWindow.SaveButton);
            this.UIMap.AcceptWindow.Acc_YesWindow.YesButton.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.AcceptWindow.Acc_YesWindow.YesButton);
            Thread.Sleep(8 * WaC);
            GetScreen("Create_Changes");
            this.UIMap.DocumentWindow.SetFocus(); // Постараемся установить фокус на окне конструктора
            Keyboard.SendKeys("{F4}", ModifierKeys.Alt);
            this.UIMap.DocumentWindow.WaitForControlNotExist(60 * WaC);
        }

        public void CreateAkt(int WaC, int lvl)
        {
            OpenReport(WaC, lvl + 1, "CreateAkt");
            this.UIMap.EndedCheckWindow.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.EndedCheckWindow.Panel.CreateAktRadioButton);
            this.UIMap.EndedCheckWindow.OKButton.WaitForControlEnabled(60 * WaC);
            Mouse.Click(this.UIMap.EndedCheckWindow.OKButton);
            Thread.Sleep(8 * WaC);
            if (this.UIMap.WarningWindow.Exists)
            {
                this.UIMap.WarningWindow.WarnNo_Window.NoButton.WaitForControlExist(60 * WaC);
                Mouse.Click(this.UIMap.WarningWindow.WarnNo_Window.NoButton);
                ConnectToWorkingARM(WaC, lvl + 1, "ASIMOB_UI", "RIO");
                OpenReport(WaC, lvl + 1, "CreateAkt");
            }
            this.UIMap.AttentionWindow.WaitForControlExist(60 * WaC);
            Mouse.Click(this.UIMap.AttentionWindow.Att_NoWindow.NoButton);
            this.UIMap.NameAktWindow.WaitForControlExist(60 * WaC);
            this.UIMap.NameAktWindow.NameAktEdit.Text = "Акт проверки (РИО)_" + DateTime.Now.ToString("ddMMyyyy_HHmmss");
            Mouse.Click(this.UIMap.NameAktWindow.OKButton);
            this.UIMap.OpenEditor.WaitForControlExist(120 * WaC);
            this.UIMap.OpenEditor.WaitForControlNotExist(120 * WaC);
            this.UIMap.ASI_Window.AktReportWindow.WaitForControlExist(120 * WaC);
            this.UIMap.CreatingWindow.WaitForControlExist(60 * WaC);
            this.UIMap.CreatingWindow.ProgressBar.WaitForControlNotExist(900 * WaC);
            Thread.Sleep(5 * WaC);
            if(this.UIMap.CreatingWindow.OKButton.Exists)
            {
                Mouse.Click(this.UIMap.CreatingWindow.OKButton);
            }
            else
            {
                Mouse.Click(Dots["OKButton"]);
            }  
            Mouse.Click(MouseButtons.Right, ModifierKeys.None, Dots["EndingPart"]);
            Thread.Sleep(1 * WaC);
            Mouse.Click(Dots["ChangeEndingPart"]);
            this.UIMap.DocumentWindow.WaitForControlExist(60 * WaC);
            Mouse.Click(Dots["SetText"]);            
            CheckLanguage(lvl + 1);
            Keyboard.SendKeys("TODAY_" + DateTime.Now.ToString("ddMMyyyy_HHmmss"));
            Mouse.Click(this.UIMap.DocumentWindow.SaveButton);
            Thread.Sleep(8 * WaC);
            GetScreen("Create_Changes");
            this.UIMap.DocumentWindow.SetFocus(); // Постараемся установить фокус на окне конструктора
            Keyboard.SendKeys("{F4}", ModifierKeys.Alt);
            this.UIMap.DocumentWindow.WaitForControlNotExist(60 * WaC);
        }

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
               SelectTU(WaC, lvl + 1);
            }
            SetAFSB(WaC, lvl + 1);
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
                Mouse.Click(this.UIMap.InfoWindow.OKWindow.OKButton);
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

        public void ConnectToWorkingARM (int WaC, int lvl, string SchemaName,string UserName)
        {
            Shutdown_asi();
            Thread.Sleep(2 * WaC);
            Clear_cache();
            StartASI(WaC, lvl + 1, SchemaName, UserName);
            this.UIMap.ASI_Window.WaitForControlExist(180 * WaC);
            Thread.Sleep(10 * WaC);
        }

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


                //PrepareAsi(WaC, 1, "ASISTA_UI");

                // DownloadSpr(WaC, 1);

                // CalcSpr(WaC, 1);
                // CreateQuestions(WaC, 1);
                // DistribTask(WaC, 1);

                // DownloadReports(WaC, 1);
                // AddFile(WaC, 1);
                // AddBorrower(WaC, 1);

                // AddCustomBorrower(WaC, 1);
                //  CheckLanguage(0);
                //AddCustomBorrowerCreateReport(WaC, 1);
                // AddCustomBorrower(WaC, 1);
                //  CreateIndRep(WaC,1);
                //  CreateAkt(WaC, 1);
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
