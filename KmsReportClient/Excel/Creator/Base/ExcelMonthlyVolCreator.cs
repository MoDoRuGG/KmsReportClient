using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KmsReportClient.External;
using KmsReportClient.Model;
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.Excel;
using KmsReportClient.Support;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Base
{
    public class ExcelMonthlyVolCreator : ExcelBaseCreator<ReportMonthlyVol>
    {

        private readonly List<ReportDictionary> _MonVolDictionaries = new List<ReportDictionary> {

            new ReportDictionary {TableName = "Стационарная помощь", StartRow = 7, EndRow = 19, Index = 1},
            new ReportDictionary {TableName = "Дневной стационар", StartRow = 21, EndRow = 33, Index = 2},
            new ReportDictionary {TableName = "АПП", StartRow = 35, EndRow = 47, Index = 3},
            new ReportDictionary {TableName = "Скорая медицинская помощь", StartRow = 49, EndRow = 61, Index = 4},
        };


        public ExcelMonthlyVolCreator(
          string filename,
          ExcelForm reportName,
          string header,
          string filialName) : base(filename, reportName, header, filialName, false)
        {
        }

        protected override void FillReport(ReportMonthlyVol report, ReportMonthlyVol yearReport)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            foreach (var themeData in report.ReportDataList.OrderBy(x => x.Theme))
            {
                var dict = _MonVolDictionaries.Single(x => x.TableName == themeData.Theme);
                var data = themeData.Data;
                switch (themeData.Theme)
                {
                    case "Стационарная помощь":
                    case "Дневной стационар":
                    case "АПП":
                    case "Скорая медицинская помощь":
                        FillTable(data, dict.StartRow, dict.EndRow, themeData.Theme);
                        break;
                }
            }
        }




        private void FillTable(ReportMonthlyVolDataDto[] data, int startRowIndex, int endRowIndex, string form)
        {
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                for (int j = 1; i <= 13; i++)
                {
                    var rowData = data?.SingleOrDefault(x => x.Code == j.ToString());
                    if (rowData != null)
                    {
                        ObjWorkSheet.Cells[i, 2] = rowData.CountSluch;
                        ObjWorkSheet.Cells[i, 3] = rowData.CountAppliedSluch;
                        ObjWorkSheet.Cells[i, 6] = rowData.CountSluchMEE;
                        ObjWorkSheet.Cells[i, 10] = rowData.CountSluchEKMP;
                    };
                }
            }
        }
    }
}
