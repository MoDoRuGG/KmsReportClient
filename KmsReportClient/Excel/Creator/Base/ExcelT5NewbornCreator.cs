using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using KmsReportClient.Support;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Base
{
    public class ExcelT5NewbornCreator : ExcelBaseCreator<ReportT5Newborn>
    {

        private EndpointSoap _client;

        private string _regionCode;

        public ExcelT5NewbornCreator(
          string filename,
          ExcelForm reportName,
          string header,
          string filialName, EndpointSoap client, string regionCode) : base(filename, reportName, header, filialName, false)
        {
            _client = client;
            _regionCode = regionCode;
        }



        protected override void FillReport(ReportT5Newborn report, ReportT5Newborn yearReport)
        {
            int sheet = 1;

            foreach (var theme in report.ReportDataList)
            {
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[sheet];

                string reportMonths = Convert.ToInt32(report.Yymm.Substring(2, 2)).ToString();
                string reportYear = report.Yymm.Substring(0, 2);

                ObjWorkSheet.Cells[2, 1] = $"Cтрахование новорожденных за {reportMonths} месяц(а/ев) 20{reportYear} года";
                ObjWorkSheet.Cells[7, 2] = FilialName;
                ObjWorkSheet.Cells[7, 3] = theme.Data.MarketShare/100;
                ObjWorkSheet.Cells[7, 4] = theme.Data.CountNewborn;
                ObjWorkSheet.Cells[7, 5] = theme.Data.CountMaterinityBills;
            }
        }
    }
}
