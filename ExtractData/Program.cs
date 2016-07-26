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
using System.Text.RegularExpressions;
using System.Threading;
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
            //FilterStore();
            GetDeliveryDays1();
        }

        public static void FilterStore()
        {
            string sourceFile = @"C:\Temp\storedest.txt";
            string destFile = @"C:\Temp\filteredStoredest.txt";
            var lines = File.ReadAllLines(sourceFile);
            var destCodes = lines[0].Split(',');
            StringBuilder sb = new StringBuilder();
            for (int i = 1; i < lines.Length; i++)
            {
                var line = lines[i];
                var rateCodes = line.Split(',');
                Dictionary<string, string> dict = new Dictionary<string, string>();
                for (int j = 1; j < rateCodes.Length; j++)
                {
                    if (!dict.ContainsKey(rateCodes[j]))
                    {
                        dict.Add(rateCodes[j], destCodes[j]);
                        rateCodes[j] = "origin" + destCodes[j];
                    }
                    else
                    {
                        rateCodes[j] = dict[rateCodes[j]];
                    }
                }
                sb.AppendLine(string.Join(",",rateCodes));
            }
            File.WriteAllText(destFile, sb.ToString());
        }

        public static void Minimize()
        {
            using (ExcelParse ep = new ExcelParse())
            {
                //string sourceFile = Console.ReadLine();
                string sourceFile = @"C:\Temp\delivery-days-ss.xlsx";
                string destFile = @"C:\Temp\filtered-delivery-days-ss.xlsx";

                ep.xlApp.DisplayAlerts = false;
                ep.LoadExcelSheet(sourceFile, 3);
                //var tasks = new List<Task>();
                Console.WriteLine(ep.xlWorkSheet.Name);
                Console.WriteLine("{0}/{1}", ep.range.Rows.Count, ep.range.Columns.Count);
                Dictionary<string, string> dict = new Dictionary<string, string>();
                for (int i = 2; i <= ep.range.Rows.Count; i++)
                {
                    for (int j = 5; j <= ep.range.Columns.Count; j++)
                    {
                        int[] cell = { i, j };
                        var r = cell[0];
                        var c = cell[1];

                        Console.WriteLine("{0}/{1}", r, c);
                        var storeCode = (string)(ep.range.Cells[r, 4] as Excel.Range).Value;
                        var desCode = (string)(ep.range.Cells[1, c] as Excel.Range).Value;
                        var rateCode = (double)(ep.range.Cells[r, c] as Excel.Range).Value;
                        var key = string.Format("{0}{1}", storeCode, rateCode);
                        if (!dict.ContainsKey(key))
                        {
                            dict.Add(key, string.Format("{0},{1}", storeCode, desCode));
                            ep.xlWorkSheet.Cells[r, c] = string.Empty;
                        }
                        else
                        {
                            ep.xlWorkSheet.Cells[r, c] = dict[key];
                        }
                    }
                    Console.WriteLine("saving file");
                    ep.xlWorkBook.SaveAs(destFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    Console.WriteLine("finish saving file");
                }
                Console.WriteLine("finishing minimizing");
            }
        }

        public static void GetDeliveryDays1()
        {
            string sourceFile = @"C:\Temp\filteredStoredest.txt";
            string destFile = @"C:\Temp\finalResult.txt";
            var lines = File.ReadAllLines(sourceFile);
            var destCodes = lines[0].Split(',');

            // lines.Length
            for (int i = 0; i < lines.Length; i++)
            {
                StringBuilder sb = new StringBuilder();
                var line = lines[i];
                var codes = line.Split(',');
                Dictionary<string, string> dict = new Dictionary<string, string>();
                var storeCode = codes[0];
                var tasks = new List<Task>();

                for (int j = 1; j < codes.Length; j++)
                {
                    if (codes[j].Contains("origin"))
                    {
                        string[] inputCodes = { storeCode , codes[j].Replace("origin", ""), i.ToString(), j.ToString() };
                        tasks.Add(Task.Factory.StartNew(ics => 
                        {
                            string[] inputs = ics as string[];
                            var result = SnatchdeliveryDays(inputs).Result;
                            if(result != null)
                            {
                                codes[int.Parse(inputs[3])] = string.Format("{0}-{1}-{2}", inputs[1], result.Item1, result.Item2);
                            }
                        }, inputCodes));
                    }
                }
                var finalTask = Task.Factory.ContinueWhenAll(tasks.ToArray(), ts => 
                {
                    sb.AppendLine(string.Join(",", codes));
                });
                finalTask.Wait();
                File.AppendAllText(destFile, sb.ToString());
            }
        }

        public static void GetDeliveryDays()
        {
            using (ExcelParse ep = new ExcelParse())
            {
                //string sourceFile = Console.ReadLine();
                string sourceFile = @"C:\Temp\delivery-days.xlsx";
                string destFile = @"C:\Temp\finalResult.xlsx";
                ep.LoadExcelSheet(sourceFile, 1);
                //ep.range.Rows.Count
                var tasks = new List<Task>();
                for (int i = 2; i <= 4; i++)
                {
                    //ep.range.Columns.Count
                    for (int j = 8; j <= 10; j++)
                    {

                        int[] cell = { i, j };

                        //tasks.Add(new Task(c =>
                        //{
                        //    var indices = c as int[];
                        //    var storeCode = (string)(ep.range.Cells[indices[0], 5] as Excel.Range).Value;
                        //    var desCode = (string)(ep.range.Cells[1, indices[1]] as Excel.Range).Value;
                        //    string[] codes = { storeCode, desCode };
                        //    var deliveryDays = SnatchdeliveryDays(codes).Result;
                        //    if (deliveryDays != null)
                        //    {
                        //        ep.xlWorkSheet.Cells[indices[0], indices[1]] = string.Format("{0},{1}", deliveryDays.Item1, deliveryDays.Item2);
                        //        //Console.WriteLine(string.Format("[{0} {1}]: {2}", i, j, deliveryDays.Item1 + "/" + deliveryDays.Item2));
                        //    }
                        //}, cell));

                        tasks.Add(Task.Factory.StartNew(c =>
                        {
                            var indices = c as int[];
                            Console.WriteLine("Get value for Cell[{0} {1}]", indices[0], indices[1]);
                            var storeCode = (string)(ep.range.Cells[indices[0], 5] as Excel.Range).Value;
                            var desCode = (string)(ep.range.Cells[1, indices[1]] as Excel.Range).Value;
                            string[] codes = { storeCode, desCode };
                            var deliveryDays = SnatchdeliveryDays(codes).Result;
                            if (deliveryDays != null)
                            {
                                ep.xlWorkSheet.Cells[indices[0], indices[1]] = string.Format("{0},{1}", deliveryDays.Item1, deliveryDays.Item2);
                                //Console.WriteLine(string.Format("[{0} {1}]: {2}", i, j, deliveryDays.Item1 + "/" + deliveryDays.Item2));
                            }
                        }, cell));

                        var finalTask = Task.Factory.ContinueWhenAll(tasks.ToArray(), snatchingTask =>
                        {
                            ep.xlWorkBook.SaveAs(destFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        });
                        finalTask.Wait();
                    }
                }
            }
        }

        private static async Task<Tuple<string,string>> SnatchdeliveryDays(string[] codes)
        {
            try
            {
                Console.WriteLine("Get delivery days for {0}/{1}/{2}/{3}", codes[0], codes[1], codes[2], codes[3]);
                httpClient = new HttpClient();
                var link = "https://www.canadapost.ca/business/tools/ds/dspage.aspx";
                var param = new NameValueCollection();
                param.Add("lang", "en");
                param.Add("language", "en");
                param.Add("service", "Small Bus");
                param.Add("srccode", codes[0]);
                param.Add("desttype", "single");
                param.Add("destcode", codes[1]);
                param.Add("ctl00$ctl00$SegmentContent$Content$SubmitBtn2", "Submit");
                var content = CreateFormEncodedData(param);
                using (var rsp = await httpClient.PostAsync(link, content))
                {
                    var page = await rsp.Content.ReadAsStringAsync();
                    var doc = new HtmlDocument();
                    doc.LoadHtml(page);
                    var PriorityDays = doc.DocumentNode.SelectSingleNode(@"//*[@id=""tableContent""]/table/tr/td[2]/table/tr[3]/td[2]").InnerText.Replace("days","").Replace("day", "").Trim();
                    var XpresspostDays = doc.DocumentNode.SelectSingleNode(@"//*[@id=""tableContent""]/table/tr/td[2]/table/tr[3]/td[3]").InnerText.Replace("days", "").Replace("day", "").Trim();
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

    //public class Example
    //{
    //    public static void Main()
    //    {
    //        string[] filenames = { "chapter1.txt", "chapter2.txt", 
    //                         "chapter3.txt", "chapter4.txt",
    //                         "chapter5.txt" };
    //        string pattern = @"\b\w+\b";
    //        var tasks = new List<Task>();
    //        int totalWords = 0;

    //        // Determine the number of words in each file.
    //        foreach (var filename in filenames)
    //            tasks.Add(Task.Factory.StartNew(fn =>
    //            {
    //                if (!File.Exists(fn.ToString()))
    //                    throw new FileNotFoundException("{0} does not exist.", filename);

    //                StreamReader sr = new StreamReader(fn.ToString());
    //                String content = sr.ReadToEnd();
    //                sr.Close();
    //                int words = Regex.Matches(content, pattern).Count;
    //                Interlocked.Add(ref totalWords, words);
    //                Console.WriteLine("{0,-25} {1,6:N0} words", fn, words);
    //            },
    //                                              filename));

    //        var finalTask = Task.Factory.ContinueWhenAll(tasks.ToArray(), wordCountTasks =>
    //        {
    //            int nSuccessfulTasks = 0;
    //            int nFailed = 0;
    //            int nFileNotFound = 0;
    //            foreach (var t in wordCountTasks)
    //            {
    //                if (t.Status == TaskStatus.RanToCompletion)
    //                    nSuccessfulTasks++;

    //                if (t.Status == TaskStatus.Faulted)
    //                {
    //                    nFailed++;
    //                    t.Exception.Handle((e) =>
    //                    {
    //                        if (e is FileNotFoundException)
    //                            nFileNotFound++;
    //                        return true;
    //                    });
    //                }
    //            }
    //            Console.WriteLine("\n{0,-25} {1,6} total words\n",
    //                              String.Format("{0} files", nSuccessfulTasks),
    //                              totalWords);
    //            if (nFailed > 0)
    //            {
    //                Console.WriteLine("{0} tasks failed for the following reasons:", nFailed);
    //                Console.WriteLine("   File not found:    {0}", nFileNotFound);
    //                if (nFailed != nFileNotFound)
    //                    Console.WriteLine("   Other:          {0}", nFailed - nFileNotFound);
    //            }
    //        });
    //        finalTask.Wait();
    //    }
    //}
}
