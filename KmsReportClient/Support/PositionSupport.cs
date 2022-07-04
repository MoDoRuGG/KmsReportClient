using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KmsReportClient.Support
{
    public static class PositionSupport
    {
        public static int GetColumn(string position)
        {
            char[] delimiterChars = { 'P', 'R', 'C' };
            string[] stringSplit = position.Split(delimiterChars);
            return Convert.ToInt32(stringSplit[2]);
        }

        public static int GetRow(string position)
        {
            char[] delimiterChars = { 'P', 'R', 'C' };
            string[] stringSplit = position.Split(delimiterChars);
            return Convert.ToInt32(stringSplit[3]);
        }

        public static int GetPage(string position)
        {
            char[] delimiterChars = { 'P', 'R', 'C' };
            string[] stringSplit = position.Split(delimiterChars);
            return Convert.ToInt32(stringSplit[1]);
        }

    }
}
