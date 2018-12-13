using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using System.Threading;
using System.Diagnostics;
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
        string WinDir = Environment.GetEnvironmentVariable("WinDir");

        [TestInitialize]
        public void TestStartup()
        {
            Console.WriteLine("Hello World!");
        }

        [TestCleanup]
        public void MyTestCleanup()
        {
            Console.WriteLine(WinDir + @"\System32\calc.exe");
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


        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
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
