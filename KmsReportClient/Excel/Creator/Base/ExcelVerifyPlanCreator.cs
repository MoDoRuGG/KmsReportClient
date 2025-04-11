using System;
using System.Collections.Generic;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.Excel;
using KmsReportClient.Support;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Base
{
    class ExcelVerifyPlanCreator : ExcelBaseCreator<ReportViolations>
    {
        public ExcelVerifyPlanCreator(
            string filename,
            ExcelForm reportName,
            string header,
            string filialName) : base(filename, reportName, header, filialName, false) { }


        private readonly List<ReportDictionary> _Dictionaries = new List<ReportDictionary> {

            new ReportDictionary {TableName = "Планы проверок", StartRow = 4, EndRow = 16, Index = 1},
        };


        protected override void FillReport(ReportViolations report, ReportViolations yearReport)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            foreach (var themeData in report.ReportDataList.OrderBy(x => x.Theme))
            {
                var dict = _Dictionaries.Single(x => x.TableName == themeData.Theme);
                var data = themeData.Data;

                for (int i = 0; i < themeData.Data.Length; i++)
                {
                    ObjWorkSheet.Cells[i + 4, 3] = data[i].Count;
                }
            }
        }
    }
}
