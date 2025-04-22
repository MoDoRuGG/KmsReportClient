using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using KmsReportClient.Support;
using System.Collections.Generic;
using System;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelConsZpzWebSite2025 : ExcelBaseCreator<ZpzForWebSite2025>
    {
        string period = "";
        public static readonly Dictionary<string, string> PeriodsWithMonth = new Dictionary<string, string>
        {
             {"03","1й квартал" },
             {"06","1е полугодие" },
             {"09","9 месяцев" },
             {"12","год" }
        };
        public ExcelConsZpzWebSite2025(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.ZpzForWebSite2025, header, filialName, true) { }

        protected override void FillReport(ZpzForWebSite2025 report, ZpzForWebSite2025 yearReport)
        {
            string year = (2000 + Convert.ToInt32(report.Yymm.Substring(0, 2))).ToString();
            string month = report.Yymm.Substring(2, 2);
            switch (month) {
                case "03":
                    period = $"1 квартал {year} года";
                    break;
                case "06":
                    period = $"1 полугодие {year} года";
                    break;
                case "09":
                    period = $"9 месяцев {year} года";
                    break;
                case "12":
                    period = $"{year} год";
                    break;
                
            
            }

            ObjWorkSheet.Cells[4, 1] = FilialName;
            ObjWorkSheet.Cells[5, 1] = period;
            int currentIndex = 10;
            foreach (var col in report.WSData)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = col.Col1;
                ObjWorkSheet.Cells[currentIndex, 2] = col.Col2;
                ObjWorkSheet.Cells[currentIndex, 3] = col.Col3;
                ObjWorkSheet.Cells[currentIndex, 4] = col.Col4;
                ObjWorkSheet.Cells[currentIndex, 5] = col.Col5;
                ObjWorkSheet.Cells[currentIndex, 6] = col.Col6;
                ObjWorkSheet.Cells[currentIndex, 8] = col.Col8;
                ObjWorkSheet.Cells[currentIndex, 9] = col.Col9;
                ObjWorkSheet.Cells[currentIndex, 10] = col.Col10;
                ObjWorkSheet.Cells[currentIndex, 11] = col.Col11;
                ObjWorkSheet.Cells[currentIndex, 12] = col.Col12;
                ObjWorkSheet.Cells[currentIndex, 13] = col.Col13;
                ObjWorkSheet.Cells[currentIndex, 14] = col.Col14;
            }

        }

    }
}
