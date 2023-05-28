using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using System.Drawing.Imaging;
using System.IO;
using System.Data;
using System.Data.OleDb;
using OpenQA.Selenium.DevTools.V111.HeapProfiler;
using System.Security.Permissions;

namespace WindowsFormsApp3
{
    public partial class Form1 : Form
    {
        public static IWebDriver driver;
        private static string baseURL;
        public static int i = 1;
        public Form1()
        {
            InitializeComponent();
            baseURL = "https://www.google.ru/search?q=%D0%BA%D0%B0%D0%BB%D1%8C%D0%BA%D1%83%D0%BB%D1%8F%D1%82%D0%BE%D1%80&newwindow=1&source=hp&ei=fr5zZJaMFMGErgT69K2YBg&iflsig=AOEireoAAAAAZHPMjk9ZqiF7dv-ZOQQE8TUM6XFC_pCK&oq=%D0%BA%D0%B0&gs_lcp=Cgdnd3Mtd2l6EAEYADILCAAQgAQQsQMQgwEyCwgAEIAEELEDEIMBMgsIABCKBRCxAxCDATILCAAQgAQQsQMQgwEyCwguEIoFELEDEIMBMgsIABCABBCxAxCDATIFCAAQgAQyEQguEIAEELEDEIMBEMcBENEDMgUIABCABDILCAAQgAQQsQMQgwE6DgguEIAEELEDEIMBENQCUABYVWDPBWgAcAB4AIABQogBgAGSAQEymAEAoAEB&sclient=gws-wiz";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("Logs\\Log_test.txt");
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("screen");
        }

        public static void ScreenShot(string x)
        {
            Screenshot ss = ((ITakesScreenshot)driver).GetScreenshot();
            string screenshot = ss.AsBase64EncodedString;
            byte[] screenshotAsByteArray = ss.AsByteArray;
            ss.SaveAsFile("screen" + x + i.ToString(), ScreenshotImageFormat.Jpeg);
            ss.ToString();
            i++;
        }
        
        public void Log(string x)
        {
            StreamWriter file = new StreamWriter("Logs\\Log_test.txt",true);
            file.WriteLine(DateTime.Now.ToString() + "|" + x);
            file.Close();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            driver = new FirefoxDriver();
            Log("Открыт браузер");
            driver.Navigate().GoToUrl(baseURL);
            driver.Manage().Timeouts();
            driver.Manage().Window.Maximize();
            //jlkklc - ввод
            System.Threading.Thread.Sleep(10000);
            IWebElement searchInput = driver.FindElement(By.ClassName("jlkklc"));
            IWebElement znak = null;
            string con = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="  +"test.xlsx"+ ";Extended Properties='Excel 12.0;HDR=NO;IMEX=1;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from [лист1$]", connection);
                using(OleDbDataReader dr = command.ExecuteReader())
                {
                    
                    
                    //XRsWPe MEdqYd - +
                    while (dr.Read())
                    {
                        driver.FindElement(By.XPath("/html/body/div[6]/div/div[10]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div[1]/div/div/div[1]/div[2]/div[2]/div")).SendKeys(dr[0].ToString());
                        Log("Введён первый операнд = "+dr[0].ToString());
                        string caseSwitch = dr[1].ToString();
                        switch (caseSwitch)
                        {
                            case "+":
                                znak = driver.FindElement(By.XPath("/html/body/div[6]/div/div[10]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div[1]/div/div/div[3]/div/table[2]/tbody/tr[5]/td[4]/div/div"));
                                Log("складывание");
                                break;
                            case "-":
                                znak = driver.FindElement(By.XPath("/html/body/div[6]/div/div[10]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div[1]/div/div/div[3]/div/table[2]/tbody/tr[4]/td[4]/div/div"));
                                Log("вычитание");
                                break;
                            case "*":
                                znak = driver.FindElement(By.XPath("/html/body/div[6]/div/div[10]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div[1]/div/div/div[3]/div/table[2]/tbody/tr[3]/td[4]/div/div"));
                                Log("умножение");
                                break;
                            case "/":
                                znak = driver.FindElement(By.XPath("/html/body/div[6]/div/div[10]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div[1]/div/div/div[3]/div/table[2]/tbody/tr[2]/td[4]/div/div"));
                                Log("деление");
                                break;
                        }
                        //searchInput.Click(); ???
                        if (znak != null)
                        {
                            znak.Click();
                        }
                        driver.FindElement(By.XPath("/html/body/div[6]/div/div[10]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div[1]/div/div/div[1]/div[2]/div[2]/div")).SendKeys(dr[2].ToString());
                        Log("Введён второй операнд = " + dr[2].ToString());

                        Log("запрос = " + driver.FindElement(By.ClassName("jlkklc")).Text.ToString());
                        driver.FindElement(By.XPath("/html/body/div[6]/div/div[10]/div/div[2]/div[2]/div/div/div[1]/div/div/div/div[1]/div/div/div[3]/div/table[2]/tbody/tr[5]/td[3]/div/div")).Click();
                        Log("результат = " + driver.FindElement(By.ClassName("jlkklc")).Text.ToString());

                        if (driver.FindElement(By.ClassName("jlkklc")).Text.ToString() == dr[3].ToString())
                        {
                            Log("Эталонный вариант = " + dr[3].ToString() + " полученный = " + driver.FindElement(By.ClassName("jlkklc")).Text.ToString() + " ответы совпали");
                        }
                        else
                        {
                            Log("Эталонный вариант = " + dr[3].ToString() + " полученный = " + driver.FindElement(By.ClassName("jlkklc")).Text.ToString() + " ответы НЕ совпали!");
                            ScreenShot("результат");
                            Log("Сделан снимок экрана №" + i);
                        }
                        
                    
                    }
                    driver.Close();
                    Log("Браузер закрыт");
                    label2.Text = "Test_Completed";
                    Log("Тест завершен");
                }
            }
        }
    }
}
