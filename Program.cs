using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Marshal = System.Runtime.InteropServices.Marshal;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Data.OleDb;
using System.Configuration;

using OpenQA.Selenium;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.PhantomJS;
using System.Diagnostics;
using System.IO;

namespace CLI_Offer
{
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        /// 
        public static IWebDriver Driver;

        public static string NetBankID = null;
        public static string NBpwd = null;
        public int noofTransactions;
        public string curtime = null;
        public int totalcount = 0;

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
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

        public void LogResults(string curTme, string User, string appln, string Env, string Toolname, int Nooftimes)
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
            else
            {
                MessageBox.Show(@"The logs were not captured because you dont have access to this folder \flsy02\gd_csl_epg_pr_sy$\Premium\T E S T I N G\DEA_DO, please send a email to get access or raise a request it for the same.", "Restricted user access");
            }

        }




        public void ReadExcelArgument(string DataFilePath, string env)
        {
            DateTime now = DateTime.Now;
            curtime = now.ToString();
            List<string> rowValue = new List<string> { };
            var ExcelFilePath = DataFilePath;
            Excel.Application xlApp = new Excel.Application(ExcelFilePath);
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ExcelFilePath);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int newIValue = 1;
                noofTransactions = rowCount - 1;
                var chromeDriverService = ChromeDriverService.CreateDefaultService();
                chromeDriverService.HideCommandPromptWindow = true;
                Driver = new ChromeDriver(chromeDriverService);
                Driver.Manage().Window.Maximize();
               
                System.Threading.Thread.Sleep(200);

                NBpwd = "12345678"; 

                //int colCount = xlRange.Columns.Count
                for (int i = 2; i <= rowCount; i++)
                {

                    try
                    {
                        NetBankID = Convert.ToString(xlRange.Cells[i, 1].Value2);
                        String Status = Convert.ToString(xlRange.Cells[i, 2].Value2);


                        
                        if ((Status == "DONE") && NetBankID != (""))
                        {
                            //Form1.InvalidArg = true;
                            //break;
                        }
                        else if (NetBankID != null && Status == null)
                        {


                            if (Form1.Env == "T2")
                            {
                                Driver.Navigate().GoToUrl("http://www.my.test2.finest.online.cba/netbank/Logon/Logon.aspx");
                            }
                            else
                            {
                                Driver.Navigate().GoToUrl("http://www.my.test5.finest.online.cba/netbank/Logon/Logon.aspx");
                            }
                            System.Threading.Thread.Sleep(200);

                            
                                Driver.FindElement(By.Id("txtMyClientNumber_field")).SendKeys(NetBankID);
                                System.Threading.Thread.Sleep(200);
                                Driver.FindElement(By.Id("txtMyPassword_field")).SendKeys(NBpwd);
                            
                            
                            Driver.FindElement(By.Id("btnLogon_field")).Click();




                            if (IsElementPresent(By.Id("btnLogon_field")))
                            {
                                xlRange.Cells[i, 2].Value2 = "Failed-Log on Error";
                                
                            }
                            else if (IsElementPresent(By.XPath("//div[contains(text (), 'Sorry')]")))
                            {

                                
                                xlRange.Cells[i, 2].Value2 = "Failed - Page not found error";
                                
                                
                            }
                            else
                            {
                                if (IsElementPresent(By.Id("ctl00_BodyPlaceHolder_CheckBoxPanelTermsAndConditionsAccept_field")))
                                {
                                    Driver.FindElement(By.Id("ctl00_BodyPlaceHolder_CheckBoxPanelTermsAndConditionsAccept_field")).Click();
                                    Driver.FindElement(By.Id("ctl00_BodyPlaceHolder_btnNext_field")).Click();

                                }
                                var WaitforSettings = new WebDriverWait(Driver, TimeSpan.FromSeconds(10));
                                IWebElement setingsElement;
                                setingsElement = WaitforSettings.Until(ExpectedConditions.ElementIsVisible(By.XPath("//a/strong[contains(.,'Settings')]")));

                                Driver.FindElement(By.XPath("//a/strong[contains(.,'Settings')]")).Click();

                                var dWait = new WebDriverWait(Driver, TimeSpan.FromSeconds(10));
                                IWebElement dynamicElement;
                                dynamicElement = dWait.Until(ExpectedConditions.ElementIsVisible(By.Id("settings-credit-card-limit-increase-invitations")));
                                Driver.FindElement(By.Id("settings-credit-card-limit-increase-invitations")).Click();
                                System.Threading.Thread.Sleep(5000);

                                if (Driver.PageSource.Contains("Sorry, you're unable to register"))
                                {
                                    xlRange.Cells[i, 2].Value2 = "Not authorised to change credit limit as NBID is not primary card holder.";
                                    xlWorkbook.Save();
                                }
                                else if (Driver.PageSource.Contains("Please select Yes or No before clicking Submit."))
                                {
                                    xlRange.Cells[i, 2].Value2 = "You ";
                                    xlWorkbook.Save();
                                }
                                else if (Driver.PageSource.Contains("Our invitation service lets you know when you're conditionally approved for a credit limit increase"))
                                {
                                    Driver.FindElement(By.Id("ctl00_BodyPlaceHolder_blkOptionPreference_rbtnOptIn_field")).Click();
                                    System.Threading.Thread.Sleep(200);
                                    Driver.FindElement(By.Id("ctl00_BodyPlaceHolder_btnSaveChanges_field")).Click();
                                    var dWait1 = new WebDriverWait(Driver, TimeSpan.FromSeconds(10));
                                    // IWebElement dynamicElement;
                                    dWait1.Until(ExpectedConditions.FrameToBeAvailableAndSwitchToIt(By.Name("ESD.CreditCards.ChangeLimit_v1.0.0.702_helix.container-1")));
                                                                       


                                    if (Driver.PageSource.Contains("You don't have any eligible credit cards."))
                                    {
                                        xlRange.Cells[i, 2].Value2 = "You don't have any eligible credit cards.";
                                        xlWorkbook.Save();
                                    }
                                    else
                                    {
                                       // string text = Driver.FindElement(By.XPath("/html/body/div[1]/div/div[3]/div/div[2]/div/div/div[2]/div/div/div/div/p/p")).GetAttribute("value");
                                        try
                                        {
                                            string ChooseText = Driver.FindElement(By.XPath("/html/body/div[1]/div/div[3]/div/div[2]/div/div/div[2]/div/div/form/div[1]/div[1]/p")).GetAttribute("value");
                                        }
                                        catch (Exception e)
                                        {
                                            Console.WriteLine(e.ToString());
                                        }
                                        try {
                                            string CurentLimit = Driver.FindElement(By.XPath("//li[@data-tab-id='creditLimitChangeTab']/button/span[2]")).GetAttribute("text");
                                        }
                                    catch (Exception e)
                                        {
                                            Console.WriteLine(e.ToString());
                                        }
                                        Driver.FindElement(By.Id("newCreditLimit")).Clear();
                                        Driver.FindElement(By.Id("newCreditLimit")).SendKeys("38000");
                                        System.Threading.Thread.Sleep(2000);

                                        if (IsElementPresent(By.Id("creditLimitIncreaseContinueActionButton")))
                                        {
                                            Driver.FindElement(By.Id("creditLimitIncreaseContinueActionButton")).Click();
                                            xlRange.Cells[i, 2].Value2 = "Done - Valid Input";
                                        }
                                        else
                                        {
                                            xlRange.Cells[i, 2].Value2 = "That's more than you're conditionally approved for. For a higher limit, you'll need to complete a more detailed application.";
                                            Driver.FindElement(By.Id("creditLimitIncreaseFullFormActionButton")).Click();
                                            
                                        }
                                        


                                        //Boolean NextbuttonDisplayed;

                                        //try
                                        //{
                                        //    //Driver.SwitchTo().Frame("ESD.CreditCards.ChangeLimit_v1.0.0.696_helix.container-2");
                                        //    IWebElement element = Driver.FindElement(By.Id("creditLimitIncreaseContinueActionButton"));
                                        //    NextbuttonDisplayed = element.Displayed;

                                        //}
                                        //catch (NoSuchElementException e)
                                        //{
                                        //    NextbuttonDisplayed = false;
                                        //}


                                        //if (NextbuttonDisplayed == true)
                                        //{
                                        //    xlRange.Cells[i, 2].Value2 = "Passed - Out of Boundary Values";
                                        //}
                                        //else
                                        //{
                                        //    xlRange.Cells[i, 2].Value2 = "Passed - Valid Input";
                                        //}
                                        //string url = Driver.Url; // get the current URL (full)
                                        //Uri currentUri = new Uri(url); // create a Uri instance of it
                                        //string baseUrl = currentUri.Authority; // just get the "base" bit of the URL
                                        //Driver.Navigate().GoToUrl(baseUrl + "/Feedback");

                                        
                                        xlWorkbook.Save();

                                    }

                                }
                                else
                                {
                                    xlRange.Cells[i, 2].Value2 = "Failed";
                                    xlWorkbook.Save();
                                }


                                newIValue++;
                            }
                        }
                        else
                        {
                            //InvalidArg = true;
                            //MessageBox.Show("Invalid Input", "Input(s) are Invalid!",
                            //MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                            break;
                        }
                    }
                    catch (NoSuchElementException errExcel)
                    {
                        Console.WriteLine(errExcel.ToString());
                        xlRange.Cells[i, 2].Value2 = "Failed - Page Error";
                        

                    }
                    catch (WebDriverTimeoutException others)
                    {
                        xlRange.Cells[i, 2].Value2 = "Failed - Web element is not available/Page technical problem.";
                        // break;
                    }
                    catch (Exception other)
                    {
                        xlRange.Cells[i, 2].Value2 = "Failed - Application Error";
                        // break;
                    }
                    xlWorkbook.Save();
                }
                Driver.Quit();
                xlWorkbook.Save();
                xlWorkbook.Close();
                xlApp.Quit();
   
}
    }

}
    
