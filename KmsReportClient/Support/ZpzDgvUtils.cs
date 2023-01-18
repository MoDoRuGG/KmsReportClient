using System;
using System.Linq;
using System.Windows.Forms;

namespace KmsReportClient.Support
{
    public static class ZpzDgvUtils
    {
        
        public static void SetRowText(object value, DataGridViewCell cell)
        {
            string cellValue = cell.Value == null ? "" : cell.Value.ToString();
            if (cellValue == "x")
                return;

            cell.Value = value;

        }

        public static string GetRowText(bool isExclusionsRow, string[] exclusionsCells, int index, decimal value)
        {
            bool inNeedExcludeCell = exclusionsCells != null ? IsNeedExcludeSum(exclusionsCells, index) : false;
            if (isExclusionsRow || inNeedExcludeCell)
            {
                return "x";
            }
            else
            {
                var filteredValue = value.ToString().Replace(",000", "");
                int oneChar = Convert.ToInt32(filteredValue.ElementAt(0));
                
                if (filteredValue.Contains("0,00"))
                {
                    return filteredValue;
                }

                return filteredValue.ToString();
            }
        }


     





        private static bool IsNeedExcludeSum(string[] exclusionsCells, int index) =>
            exclusionsCells?.Contains(index.ToString()) ?? false;
    }
}
