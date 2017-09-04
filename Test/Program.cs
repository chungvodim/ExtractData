using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            //StringBuilder sb = new StringBuilder();
            //sb.AppendLine("SET IDENTITY_INSERT [dbo].[DashboardStatistics] ON ");
            //var lines = File.ReadAllLines("D:\\Tasks\\Belgium-Mockups\\DashboardStatictis.csv");
            //int TotalPageViews = 0;
            //int TotalPhoneClicks = 0;
            //int TotalLeads = 0;
            //for (int i = 1; i < lines.Length; i++)
            //{
            //    var line = lines[i];
            //    var stuffs = line.Split(',');
            //    TotalPageViews += Convert.ToInt32(stuffs[4]);
            //    TotalPhoneClicks += Convert.ToInt32(stuffs[5]);
            //    TotalLeads += Convert.ToInt32(stuffs[6]);
            //    var command = "INSERT [dbo].[DashboardStatistics]([DashboardStatisticID],[CompanyID],[BranchID],[ListingID],[DateActivity],[NumPageViews],[NumPhoneClicks],[NumLeads],[NumChatMessages],[NumEmails],[TotalPageViews],[TotalPhoneClicks],[TotalLeads],[TotalChatMessages],[TotalEmails]) VALUES (";
            //    command += stuffs[0] + ",";//[DashboardStatisticID]
            //    command += stuffs[1] + ",";//[CompanyID]
            //    command += stuffs[2] + ",";//[BranchID]
            //    command += "0,";//[ListingID]
            //    command += "'" + stuffs[3] + "',";//[DateActivity]
            //    command += stuffs[4] + ",";//[NumPageViews]
            //    command += stuffs[5] + ",";//[NumPhoneClicks]
            //    command += stuffs[6] + ",";//[NumLeads]
            //    command += "0,";//[NumChatMessages]
            //    command += "0,";//[NumEmails]
            //    command += TotalPageViews + ",";//[TotalPageViews]
            //    command += TotalPhoneClicks + ",";//[TotalPhoneClicks]
            //    command += TotalLeads + ",";//[TotalLeads]
            //    command += "0,";//[TotalChatMessages]
            //    command += "0)";//[TotalEmails]
            //    sb.AppendLine(command);
            //}
            //sb.AppendLine("SET IDENTITY_INSERT [dbo].[DashboardStatistics] OFF ");
            //File.WriteAllText("D:\\Tasks\\Belgium-Mockups\\DashboardStatictis.sql", sb.ToString());
            Console.WriteLine(CalculateNextMonth());
        }

        public static string NormalizePhoneNumber(string phoneNumber)
        {
            Regex rgx = new Regex("[^0-9]");
            phoneNumber = rgx.Replace(phoneNumber, "");
            return phoneNumber;
        }

        public static DateTime CalculateNextMonth()
        {
            DateTime date = new DateTime(2017,2,28);
            var last = date;
            var nextMonth = last.AddMonths(1);
            var daysInNextMonth = DateTime.DaysInMonth(nextMonth.Year, nextMonth.Month);
            var next = last.AddDays(daysInNextMonth);
            var firstOfNextMonth = new DateTime(nextMonth.Year, nextMonth.Month, 1);
            next = next < firstOfNextMonth ? firstOfNextMonth : next;
            return nextMonth;
        }
    }
}
