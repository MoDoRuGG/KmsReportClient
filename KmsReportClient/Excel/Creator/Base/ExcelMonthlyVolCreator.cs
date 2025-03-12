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
                FillTable(data, dict.StartRow, dict.EndRow, themeData.Theme);
            }
        }




        private void FillTable(ReportMonthlyVolDataDto[] data, int startRowIndex, int endRowIndex, string form)
        {
            int j = startRowIndex;
            if (data != null)
            {
                foreach (var row in data)
                {
                    ObjWorkSheet.Cells[j, 2] = row.CountSluch;
                    ObjWorkSheet.Cells[j, 3] = row.CountAppliedSluch;
                    ObjWorkSheet.Cells[j, 6] = row.CountSluchMEE;
                    ObjWorkSheet.Cells[j++, 10] = row.CountSluchEKMP;
                    if (j == endRowIndex)
                    break;
                }
            };
        }
    }
}
