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
    class ExcelZpz2025LethalCreator : ExcelBaseCreator<ReportZpz2025>
    {
        private readonly List<ReportDictionary> _zpzDictionaries = new List<ReportDictionary> {
            new ReportDictionary {TableName = "Таблица 1Л", StartRow = 5, EndRow = 28, Index = 1},
            new ReportDictionary {TableName = "Таблица 2Л", StartRow = 5, EndRow = 30, Index = 2},
        };

        public ExcelZpz2025LethalCreator(
            string filename,
            ExcelForm reportName,
            string header,
            string filialName) : base(filename, reportName, header, filialName, false) { }

        protected override void FillReport(ReportZpz2025 report, ReportZpz2025 yearReport)
        {
            string reportMonths = YymmUtils.GetMonth(report.Yymm.Substring(2, 2));
            string reportYear = report.Yymm.Substring(0, 2);

            ObjWorkSheet.Cells[3, 1] = $"за {reportMonths} 20{reportYear} года";
            ObjWorkSheet.Cells[4, 1] = FilialName;

            foreach (var themeData in report.ReportDataList.OrderBy(x => x.Theme))
            {
                var dict = _zpzDictionaries.Single(x => x.TableName == themeData.Theme);
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[dict.Index];
                var data = themeData.Data;
                switch (themeData.Theme)
                {
                    case "Таблица 1Л":
                    case "Таблица 2Л":
                        FillTableLetal(data, dict.StartRow, dict.EndRow, themeData.Theme);
                        break;
                }
            }


            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[2];
            FinishZpz();
        }


        private void FillTableLetal(ReportZpz2025DataDto[] data, int startRowIndex, int endRowIndex, string form)
        {
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                string rowNum = ObjWorkSheet.Cells[i, 7].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var rowData = data?.SingleOrDefault(x => x.Code == rowNum);
                    if (rowData != null)
                    {
                        if (ObjWorkSheet.Cells[i, 8].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 8] = rowData.CountAmbulatory;
                        }

                        if (ObjWorkSheet.Cells[i, 9].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 9] = rowData.CountStac;
                        }

                        if (ObjWorkSheet.Cells[i, 10].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 10] = rowData.CountDs;
                        }

                        if (ObjWorkSheet.Cells[i, 11].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 11] = rowData.CountOutOfSmoAnother;
                        }

                        if (ObjWorkSheet.Cells[i, 12].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 12] = rowData.CountSmo;
                        }

                    }
                }
            }
        }

        private void FinishZpz()
        {
            ObjWorkSheet.Cells[41, 3] = CurrentUser.Director;
            ObjWorkSheet.Cells[44, 1] = "Дата: " + DateTime.Today.ToShortDateString();
            if (!string.IsNullOrEmpty(CurrentUser.DirectorPhone))
            {
                var code = GetPhoneCode(CurrentUser.DirectorPhone);
                var number = GetPhoneNumber(CurrentUser.DirectorPhone);
                ObjWorkSheet.Cells[44, 4] = $"+7 ({code}) {number}";
            }

            ObjWorkSheet.Cells[47, 3] = CurrentUser.UserName;
            ObjWorkSheet.Cells[50, 1] = CurrentUser.Email ?? "";
            if (!string.IsNullOrEmpty(CurrentUser.Phone))
            {
                var code = GetPhoneCode(CurrentUser.Phone);
                var number = GetPhoneNumber(CurrentUser.Phone);
                ObjWorkSheet.Cells[50, 4] = $"+7 ({code}) {number}";
            }
        }

    }
}
