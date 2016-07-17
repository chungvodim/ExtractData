using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.PhantomJS;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExtractData
{
    class Program
    {
        static void Main(string[] args)
        {
            IWebDriver webDriver = new ChromeDriver();
            WebDriverWait waitDriver = new WebDriverWait(webDriver, new TimeSpan(0, 3, 0));

            webDriver.Navigate().GoToUrl("https://www.canadapost.ca/tools/pg/prices/FSA_RateCodeLookup-e.asp");
            var postalCodeTB = waitDriver.Until(ExpectedConditions.ElementIsVisible(By.Id("userFSA")));
            var findButton = waitDriver.Until(ExpectedConditions.ElementToBeClickable(By.Id("buttonFSA")));
            //var host = "https://www.canadapost.ca/tools/pg/prices";
            var filePath = "C:\\Temp\\data\\";

            HttpClient httpClient = new HttpClient();
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            string str;
            int rCnt = 0;
            int cCnt = 0;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open("C:\\Temp\\ratecode-look-up-table.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            Console.WriteLine(xlWorkSheet.Name);
            #region download
            for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
            {
                str = (string)(range.Cells[rCnt, 5] as Excel.Range).Value2;
                Console.WriteLine(str);
                if (!string.IsNullOrWhiteSpace(str))
                {
                    webDriver.Navigate().Refresh();
                    postalCodeTB = waitDriver.Until(ExpectedConditions.ElementIsVisible(By.Id("userFSA")));
                    findButton = waitDriver.Until(ExpectedConditions.ElementToBeClickable(By.Id("buttonFSA")));
                    postalCodeTB.SendKeys(str);
                    findButton.Click();
                    var csvanchor = waitDriver.Until(ExpectedConditions.ElementToBeClickable(By.Id("RateCodeCSV")));
                    var csvlink = csvanchor.GetAttribute("href");
                    Console.WriteLine(csvlink);
                    //using (var rsp = httpClient.GetAsync("https://www.canadapost.ca/tools/pg/prices/E50-e.pdf").Result)
                    //{
                    //    var page = rsp.Content.ReadAsStringAsync().Result;
                    //}
                    using (var rsp = httpClient.GetAsync(csvlink).Result)
                    {
                        using(Stream contentStream = rsp.Content.ReadAsStreamAsync().Result)
                        {
                            using (var stream = new FileStream(filePath + csvlink.Replace("https://www.canadapost.ca/tools/pg/prices/csv/",""), FileMode.Create, FileAccess.Write, FileShare.None, 1024, true))
                            {
                                contentStream.CopyToAsync(stream);
                            }
                        }
                    }
                }

                //for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                //{
                //    str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                //    Console.WriteLine(str);
                //}
            }
            #endregion
            //Excel.Sheets worksheets = xlWorkBook.Worksheets;
            //var xlNewSheet = (Excel.Worksheet)worksheets.Add(worksheets[1], Type.Missing, Type.Missing, Type.Missing);
            //xlNewSheet.Name = "newsheet";

            xlWorkBook.Close(false, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine(ex.Message);
                //MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
