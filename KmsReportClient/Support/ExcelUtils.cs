using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KmsReportClient.Support
{
    public static class ExcelUtils
    {
        public static void SetCellValue(dynamic cell, object value)
        {
            if (cell.Text == "х")
                return;

            cell = value;
        }
    }
}
