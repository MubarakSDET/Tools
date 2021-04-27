//'============================================================================
//'Addition of BPAYaddress book entries for NB IDs
//'Created By Arun G & Mubarak
//'Created on 22-Feb-2017
//'=============================================================================

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenQA.Selenium;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System.Diagnostics;
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;

namespace AddingBPAYAddress
{
    public class CommandPromt  
    {
        public static IWebDriver Driver;
        public static String NetBankID ;
        public static String NBpwd;
        public static String DataFilePath;
        public static String DataFileName;
        public static String EnvironmentName = "T5";
        public static String exitFlag = null;

        public static String BaseUrl = null;
        public static String SmsUrl= null;
        public static String TransferPayUrl = null;

        int count = 0;
        public static string Today = DateTime.Today.ToString();
        public static string Time = DateTime.Now.ToString("HH:mm:ss tt");
        public static string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
        public static string Toolname = "AddingBPAYAddress";
        public static string Appln = "NETBANK";
        public static int noofTransactions;
        public static string curtime = null;
        public int totalcount = 0;
        public static string Env = null;

        public static String BaseUrlT5="https://www.my.test5.finest.online.cba/netbank/Logon/Logon.aspx";
        public static String SmsUrlT5= "http://sms.test0.caas.cba/SmsGatewayStub/testpage.aspx";
        public static String TransferPayUrlT5= "https://www1.my.test5.finest.online.cba/netbank/PaymentHub/AddressBook/AddressBook.aspx?";

        public static String BaseUrlT2 = "http://www.my.test2.finest.online.cba/netbank/Logon/Logon.aspx";
        public static String SmsUrlT2 = "http://smsgw.test2.caas.cba/SMSGatewayStub/TestPage.aspx";
        public static String TransferPayUrlT2 = "https://www1.my.test2.finest.online.cba/netbank/PaymentHub/AddressBook/AddressBook.aspx?";

        public static String BaseUrlT0 = "https://www.my.test0.finest.online.cba/netbank/Logon/Logon.aspx";
        public static String SmsUrlT0 = "http://smsgw.test4.caas.cba/SMSGatewayStub/TestPage.aspx";
        public static String TransferPayUrlT0 = "https://www1.my.test0.finest.online.cba/netbank/PaymentHub/AddressBook/AddressBook.aspx?";

        public static String BaseUrlT4 = "https://www.my.staging.commbank.com.au/netbank/Logon/Logon.aspx";
        public static String SmsUrlT4 = "https://sms.staging.caas.cba/SmsGatewayService.CAAS/SmsOutbound.asmx";
        public static String TransferPayUrlT4 = "https://www1.my.staging.commbank.com.au/netbank/PaymentHub/AddressBook/AddressBook.aspx?";

       
        static void Main(string[] args)
        {
            CommandPromt cmdObj = new CommandPromt();
            try { 
            Console.WriteLine("***************************** Welcome to Tools DEA *****************************");
            Console.WriteLine();
            Console.WriteLine("Please enter Netbank ID 'OR' press ENTER key to use T5 Environment Netbank ID '75054985'");
            Console.WriteLine();
            NetBankID = Console.ReadLine(); // Get string from user
            if (NetBankID == "") { NetBankID = "75054985"; }
            Console.WriteLine();
            //Console.WriteLine(NetBankID);

            Console.WriteLine("Please enter Netbank password 'OR' press ENTER key to use default Netbank password '12345678'");
            Console.WriteLine();
            NBpwd = Console.ReadLine();
            if (NBpwd == "") { NBpwd = "12345678"; }
            Console.WriteLine();
            //Console.WriteLine(NBpwd);


            Console.WriteLine(@"Please enter the Excel filename which is present in C:\Temp 'OR' Press ENTER key to use 'AddingBPAYAddressData.xlsx' file\n");
            Console.WriteLine();
            DataFileName = Console.ReadLine(); // Get string from user
            if (DataFileName == "") {
                DataFileName = "AddingBPAYAddressData.xlsx";
                DataFilePath = @"C:\Temp\" + DataFileName;
                //Console.WriteLine("No Filename was entered, Exiting the Tool"); exitFlag = "True"; 
            } 
                else { DataFilePath = @"C:\Temp\" + DataFileName; }
            Console.WriteLine();
            Console.WriteLine("\nPlease enter Tool target environment:\n-------------------------------------------------------------------------------\nEnter '1' for T0/SY85/TB1---ENV09 Environment\nEnter '2' for T2/SY90/UB1---ENV07 Environment\nEnter '3' for T5/SY82/UB3---ENV06 Environment\nEnter '4' for STAGING/SY88/SB1 Environment\n\n\tOR\n\nPress ENTER key to take T5 environment as default input\n");
            EnvironmentName = Console.ReadLine();
            if (EnvironmentName == "") { EnvironmentName = "3"; }

            switch (EnvironmentName)
            {
                case "3":
                    BaseUrl = BaseUrlT5;
                    SmsUrl = SmsUrlT5;
                    Env = "T5";
                    TransferPayUrl = TransferPayUrlT5;
                    NetbankLogin(); 
                    break;
                
                case "2":
                    BaseUrl = BaseUrlT2;
                    SmsUrl = SmsUrlT2;
                    Env = "T2";
                    TransferPayUrl = TransferPayUrlT2;
                    NetbankLogin(); 
                    break;

                case "1":
                    BaseUrl = BaseUrlT0;
                    SmsUrl = SmsUrlT0;
                    Env = "T0";
                    TransferPayUrl = TransferPayUrlT0;
                    NetbankLogin(); 
                    break;

                case "4":
                    BaseUrl = BaseUrlT4;
                    SmsUrl = SmsUrlT4;
                    Env = "Staging";
                    TransferPayUrl = TransferPayUrlT4;
                    NetbankLogin(); 
                    break;

                default:
                    Console.WriteLine("\nYou have entered Invalid parameters\n");
                    Console.WriteLine("Press ENTER to Quit this window\n");
                    String exitCmd = Console.ReadLine();
                    break;
               
            }

            
             
                
            }
            catch (System.IndexOutOfRangeException)
            {
                Console.WriteLine("\nYou have entered Invalid parameters, See usage with -h ,Press Enter to continue");
                Console.ReadLine();
            }
        }

        public static bool IsElementPresent(By by)
        {
            try
            {
                Driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }



        public static void NetbankLogin()
        {

            
            var chromeDriverService = ChromeDriverService.CreateDefaultService();
            chromeDriverService.HideCommandPromptWindow = true;
            Driver = new ChromeDriver(chromeDriverService);
            //Driver.Manage().Window.Maximize();
            Driver.Navigate().GoToUrl(BaseUrl);
            
            System.Threading.Thread.Sleep(200);
            Driver.FindElement(By.Id("txtMyClientNumber_field")).SendKeys(NetBankID);
            System.Threading.Thread.Sleep(200);
            Driver.FindElement(By.Id("txtMyPassword_field")).SendKeys(NBpwd);
            Driver.FindElement(By.Id("btnLogon_field")).Click();
            System.Threading.Thread.Sleep(1000);
            if (IsElementPresent(By.ClassName("PageTitle")))
            {
                new SelectElement(Driver.FindElement(By.Id("ctl00_BodyPlaceHolder_ddlQuestion1_field"))).SelectByText("What is the name of the first company that employed you?");
                Driver.FindElement(By.Id("ctl00_BodyPlaceHolder_txtAnswer1_field")).SendKeys("CBA");
                System.Threading.Thread.Sleep(1000);
                new SelectElement(Driver.FindElement(By.Id("ctl00_BodyPlaceHolder_ddlQuestion2_field"))).SelectByText("What country did you first visit on your first overseas trip?");
                Driver.FindElement(By.Id("ctl00_BodyPlaceHolder_txtAnswer2_field")).SendKeys("Australia");
                System.Threading.Thread.Sleep(1000);
                Driver.FindElement(By.Id("ctl00_BodyPlaceHolder_btnNext_field")).Click();
            }
            else
            {
               // Console.WriteLine("Element is not displayed_QA");
            }

            try
            {
                ReadExcelArgument();
                Driver.Close();

                Console.WriteLine("Press ENTER to Quit this window");

                String exitCmd = Console.ReadLine();

                Driver.Quit();

                Process[] chromeDriverProcesses = Process.GetProcessesByName("chromedriver");

                foreach (var chromeDriverProcess in chromeDriverProcesses)
                {
                    chromeDriverProcess.Kill();
                }
            }
            catch (Exception)
            {
                Driver.Close();
                Driver.Quit();

                Process[] chromeDriverProcesses = Process.GetProcessesByName("chromedriver");

                foreach (var chromeDriverProcess in chromeDriverProcesses)
                {
                    chromeDriverProcess.Kill();
                }
            }
            
            
        }


        public static void CreateBPAYAddress(String Data, String Data1, String Data2)
        {

            Driver.Navigate().GoToUrl(TransferPayUrl);
            Driver.FindElement(By.Id("ctl00_BodyPlaceHolder_btnAddPayeeNew")).Click();
            System.Threading.Thread.Sleep(1000);
            Driver.FindElement(By.Id("ctl00_BodyPlaceHolder_txtPayeeName_field")).SendKeys(Data);
            Driver.FindElement(By.Id("ctl00_BodyPlaceHolder_switchPayeeType_field_1_label")).Click();
            Driver.FindElement(By.Id("ctl00_BodyPlaceHolder_txtBillerCode_field")).SendKeys(Data1);
            Driver.FindElement(By.Id("ctl00_BodyPlaceHolder_txtReferenceNo_field")).SendKeys(Data2);
            System.Threading.Thread.Sleep(3000);
            Driver.FindElement(By.Id("ctl00_BodyPlaceHolder_txtBillName_field")).SendKeys("Adding BPAY Address");

            System.Threading.Thread.Sleep(2000);
            Driver.FindElement(By.CssSelector("#ctl00_BodyPlaceHolder_btnSave > strong")).Click();
            System.Threading.Thread.Sleep(3000);

            Driver.FindElement(By.Id("ctl00_BodyPlaceHolder_stepUp_ctl08_field")).Click();
            System.Threading.Thread.Sleep(2000);

            var chromeDriverService_new = ChromeDriverService.CreateDefaultService();
            chromeDriverService_new.HideCommandPromptWindow = true;
            IWebDriver NewBrowser = new ChromeDriver(chromeDriverService_new);
            //NewBrowser.Manage().Window.Maximize();
            NewBrowser.Navigate().GoToUrl(SmsUrl);

            System.Threading.Thread.Sleep(500);
            NewBrowser.FindElement(By.Id("txtLoginId")).SendKeys(NetBankID);
            System.Threading.Thread.Sleep(500);
            NewBrowser.FindElement(By.Id("btnRetrieve")).Click();
            System.Threading.Thread.Sleep(500);
            String codeTemp = NewBrowser.FindElement(By.XPath("/html/body/form/table/tbody/tr[2]/td[4]")).Text;
            String OTP = codeTemp.Split(' ').Last();
            System.Threading.Thread.Sleep(2000);
            NewBrowser.Close();
            NewBrowser.Quit();


            System.Threading.Thread.Sleep(200);
            Driver.FindElement(By.Id("ctl00_BodyPlaceHolder_stepUp_ctl09_field")).SendKeys(OTP);
            Driver.FindElement(By.Id("ctl00_BodyPlaceHolder_btnConfirm")).Click();
            


        }
        public static void ExitBrowser()
        {
            //Driver.FindElement(By.Id("ctl00_HeaderControl_logOffLink")).Click();
            Driver.Close();
            Driver.Quit();
            
        }

        public static void ReadExcelArgument()
        {

            List<string> rowValue = new List<string> { };
            var ExcelFilePath = DataFilePath;               

            Excel.Application xlApp = new Excel.Application(ExcelFilePath);
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ExcelFilePath);
            Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count; //UsedRange.SpecialCells(xlCellTypeLastCell).Row
            int lastRow = xlWorksheet.Cells[rowCount, 1].End(Excel.XlDirection.xlUp).Row;
            int newIValue = 1;
            //int usedRows = xlWorksheet.UsedRange.Rows.End(xlUp).Row;
                //Range("A" & Sheets("EFT").Rows.Count).End(xlUp).Row;
            noofTransactions = lastRow-1;
            for (int i = 2; i <= lastRow; i++)
            {
                DateTime now = DateTime.Now;
                curtime = now.ToString();

                  String ReadData1 = Convert.ToString(xlRange.Cells[i, 1].Value2);
                  String ReadData2 = Convert.ToString(xlRange.Cells[i, 2].Value2);
                  String ReadData3 = Convert.ToString(xlRange.Cells[i, 3].Value2);
                  String ReadStatus = Convert.ToString(xlRange.Cells[i, 4].Value2);
                  //Console.WriteLine(ReadStatus);
                  

                    if ((ReadStatus == "DONE") && ReadData1!=("")) 
                    {
                        //break;
                    }
                    else if (ReadData1 != null && ReadStatus == null)
                    {
                        
                        CreateBPAYAddress(ReadData1, ReadData2, ReadData3);
                        ProgressBar(newIValue, lastRow);
                        xlRange.Cells[i, 4].Value2 = "DONE";
                        
                        newIValue++;
                    }
                    else
                    {
                        Console.WriteLine("\nNull value occured in Excel data sheet");
                        break;
                    }
                
                xlWorkbook.Save();
            }
            LogResults(curtime, userName, Appln, Env, Toolname, noofTransactions);
            //xlWorkbook.Save();
            xlWorkbook.Close();
            xlApp.Quit();
            //Console.Read();
        }

        public static void LogResults(string curTme, string User, string appln, string Env, string Toolname, int Nooftimes)
        {
            // string filepath = @"U:\Visual Studio 2013\Projects\HL Top-Up\Logs\NetbankAutomationLogResults.txt";
            string filepath = @"\\flsy02\gd_csl_epg_pr_sy$\Premium\T E S T I N G\DEA_DO\TOOL USUAGE STATUS\AUTOMATION LOGS\AutomationLogResults.txt";

            if (File.Exists(filepath))
            {

                using (StreamWriter writer = new StreamWriter(filepath, true))
                {

                    writer.Write(curTme + ", " + User + ", " + appln + ", " + Env + ", " + Toolname + ", " + Nooftimes + Environment.NewLine);

                }
            }

        }
        public static void ProgressBar(int progress, int tot)
        {
            //draw empty progress bar
            Console.CursorLeft = 0;
            Console.Write("["); //start
            Console.CursorLeft = 32;
            Console.Write("]"); //end
            Console.CursorLeft = 1;
            float onechunk = 30.0f / tot;

            //draw filled part
            int position = 1;
            for (int i = 0; i < onechunk * progress; i++)
            {
                
                Console.BackgroundColor = ConsoleColor.Green;
                Console.CursorLeft = position++;
                Console.Write(" ");
            }
            
            //draw unfilled part
            for (int i = position; i <= 31; i++)
            {
                
                Console.BackgroundColor = ConsoleColor.Gray;
                Console.CursorLeft = position++;
                Console.Write(" ");
            }

            //draw totals
            
            Console.CursorLeft = 35;
            Console.BackgroundColor = ConsoleColor.Black;
            Console.Write(progress.ToString() + " Record(s) Completed  "); //blanks at the end remove any excess
        }
    
    }

   
}