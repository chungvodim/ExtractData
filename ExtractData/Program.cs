using HtmlAgilityPack;
using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.PhantomJS;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExtractData
{
    class Program
    {
        private static IWebDriver webDriver;
        private static WebDriverWait waitDriver;
        private static HttpClient httpClient;
        
        static void Main(string[] args)
        {
            GetDeliveryDays();
        }

        public static void GetDeliveryDays()
        {
            using (ExcelParse ep = new ExcelParse())
            {
                //string sourceFile = Console.ReadLine();
                string sourceFile = @"C:\Temp\delivery-days.xlsx";
                string destFile = @"C:\Temp\delivery-days-result.xlsx";
                ep.LoadExcelSheet(sourceFile, 1);
                for (int i = 2; i <= 10; i++)
                {
                    for (int j = 8; j <= 10; j++)
                    {
                        var storeCode = (string)(ep.range.Cells[i, 5] as Excel.Range).Value;
                        var desCode = (string)(ep.range.Cells[1, j] as Excel.Range).Value;

                        var deliveryDays = SnatchdeliveryDays(storeCode, desCode).Result;
                        if (deliveryDays != null)
                        {
                            ep.xlWorkSheet.Cells[i, j] = string.Format("{0},{1}", deliveryDays.Item1, deliveryDays.Item2);
                            //Console.WriteLine(string.Format("[{0} {1}]: {2}", i, j, deliveryDays.Item1 + "/" + deliveryDays.Item2));
                        }
                        //ep.xlWorkSheet.Cells[i, j] = string.Format("{0},{1}", i, j);
                        Console.WriteLine("Get Next Value");
                    }
                }
                ep.xlWorkBook.SaveAs(destFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
        }

        private static async Task<Tuple<string,string>> SnatchdeliveryDays(string storeCode, string desCode)
        {
            try
            {
                httpClient = new HttpClient();
                var link = "https://www.canadapost.ca/business/tools/ds/dspage.aspx";
                var param = new NameValueCollection();
                param.Add("lang", "en");
                param.Add("language", "en");
                param.Add("service", "Small Bus");
                param.Add("srccode", storeCode);
                param.Add("desttype", "single");
                param.Add("destcode", desCode);
                param.Add("ctl00$ctl00$SegmentContent$Content$SubmitBtn2", "Submit");
                var content = CreateFormEncodedData(param);
                using (var rsp = await httpClient.PostAsync(link, content))
                {
                    var page = await rsp.Content.ReadAsStringAsync();
                    var doc = new HtmlDocument();
                    doc.LoadHtml(page);
                    var PriorityDays = doc.DocumentNode.SelectSingleNode(@"//*[@id=""tableContent""]/table/tr/td[2]/table/tr[3]/td[2]").InnerText.Replace("days","").Trim();
                    var XpresspostDays = doc.DocumentNode.SelectSingleNode(@"//*[@id=""tableContent""]/table/tr/td[2]/table/tr[3]/td[3]").InnerText.Replace("days", "").Trim();
                    return new Tuple<string, string>(PriorityDays, XpresspostDays);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }
        public static void GetPostalCode() 
        {
            httpClient = new HttpClient();
            webDriver = new ChromeDriver();
            waitDriver = new WebDriverWait(webDriver, new TimeSpan(0, 3, 0));
            webDriver.Navigate().GoToUrl("https://www.canadapost.ca/tools/pg/prices/FSA_RateCodeLookup-e.asp");
            var postalCodeTB = waitDriver.Until(ExpectedConditions.ElementIsVisible(By.Id("userFSA")));
            var findButton = waitDriver.Until(ExpectedConditions.ElementToBeClickable(By.Id("buttonFSA")));
            var filePath = "C:\\Temp\\data\\";

            string str;
            int rCnt = 0;

            using (ExcelParse ep = new ExcelParse())
            {
                string sourceFile = Console.ReadLine();
                ep.LoadExcelSheet(sourceFile, 1);

                for (rCnt = 2; rCnt <= ep.range.Rows.Count; rCnt++)
                {
                    str = (string)(ep.range.Cells[rCnt, 5] as Excel.Range).Value2;
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
                        using (var rsp = httpClient.GetAsync(csvlink).Result)
                        {
                            using (Stream contentStream = rsp.Content.ReadAsStreamAsync().Result)
                            {
                                using (var stream = new FileStream(filePath + csvlink.Replace("https://www.canadapost.ca/tools/pg/prices/csv/", ""), FileMode.Create, FileAccess.Write, FileShare.None, 1024, true))
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
            }
        }

        public static StringContent CreateFormEncodedData(NameValueCollection keyValues)
        {
            return CreateFormEncodedData(keyValues, Encoding.UTF8);
        }

        public static StringContent CreateFormEncodedData(NameValueCollection keyValues, Encoding encoding)
        {
            return CreateFormEncodedData(ToQueryString(keyValues, encoding));
        }

        public static StringContent CreateFormEncodedData(string queryString)
        {
            return CreateFormEncodedData(queryString, Encoding.UTF8);
        }

        public static StringContent CreateFormEncodedData(string queryString, Encoding encoding)
        {
            return new StringContent(queryString, encoding, "application/x-www-form-urlencoded");
        }

        public static string ToQueryString(NameValueCollection collection, Encoding encoding)
        {
            var array = (from key in collection.AllKeys
                         from value in collection.GetValues(key)
                         select string.Format("{0}={1}", HttpUtility.UrlEncode(key), HttpUtility.UrlEncode(value, encoding))).ToArray();
            return string.Join("&", array);
        }
    }

}
