using System;
using System.Collections.Generic;
using KmsReportClient.Global;

namespace KmsReportClient.Support
{
    static class YymmUtils
    {

        

        public static string ConvertPeriodToYymm(string yymm)
        {
            string[] parts = yymm.Split(' ');
            string numMonth = (Array.IndexOf(GlobalConst.Months, parts[0]) + 1).ToString();
            return parts[1].Substring(2, 2) +
                   (numMonth.Length == 1 ? "0" + numMonth : numMonth);
        }

        public static DateTime ConvertYymmToDate(string yymm)
        {
            int year = 2000 + Convert.ToInt32(yymm.Substring(0, 2));
            int month = Convert.ToInt32(yymm.Substring(2, 2));
            return new DateTime(year, month, 1);
        }

        public static string GetMonth(string mm)
        {
            int month = Convert.ToInt32(mm);
            return GlobalConst.Months[month - 1];
        }

        public static string GetYymmFromInt(object year, object month)
        {
            int convertedYear = Convert.ToInt32(year);
            int convertedMonth = Convert.ToInt32(month);

            string yymm = Convert.ToString(convertedYear % 2000);
            yymm += convertedMonth < 10 ? $"0{convertedMonth}" : convertedMonth.ToString();
            return yymm;
        }

        public static List<KeyValuePair<int, string>> GetMonths()
        {
            var monthsSpr = new List<KeyValuePair<int, string>>();
            for (int i = 1; i <= 12; i++)
            {
                int key = i;
                string month = new DateTime(DateTime.Today.Year, i, 1).ToString("MMMM");
                monthsSpr.Add(new KeyValuePair<int, string>(key, month));
            }

            return monthsSpr;
        }
    }
}
