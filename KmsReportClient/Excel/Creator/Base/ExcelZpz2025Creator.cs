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
    class ExcelZpz2025Creator : ExcelBaseCreator<ReportZpz2025>
    {
        private readonly List<ReportDictionary> _zpzDictionaries = new List<ReportDictionary> {

            new ReportDictionary {TableName = "Таблица 1", StartRow = 12, EndRow = 99, Index = 1},
            new ReportDictionary {TableName = "Таблица 2", StartRow = 7, EndRow = 27, Index = 2},
            new ReportDictionary {TableName = "Таблица 3", StartRow = 7, EndRow = 51, Index = 3},
            new ReportDictionary {TableName = "Таблица 4", StartRow = 7, EndRow = 13, Index = 4},
        };

        public ExcelZpz2025Creator(
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
                    case "Таблица 1":
                        FillTable1(data, dict.StartRow, dict.EndRow, themeData.Theme);
                        break;
                  
                    case "Таблица 4":
                        FillTable4(data, dict.StartRow, dict.EndRow, themeData.Theme);
                        break;
                    case "Таблица 2":
                    case "Таблица 3":
                        FillTable2(data, dict.StartRow, dict.EndRow);
                        break;
                }
            }
        }


        private void FillTable4(ReportZpz2025DataDto[] data, int startRowIndex, int endRowIndex, string form)
        {
            var columnIndex = form switch
            {
                "Таблица 4" => 5,
            };
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                string rowNum = ObjWorkSheet.Cells[i, 2].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var rowData = data?.SingleOrDefault(x => x.Code == rowNum);
                    if (rowData != null)
                    {
                        ObjWorkSheet.Cells[i, columnIndex] = rowData.CountSmo;
                    }
                }
            }
        }

        private void FillTable1(ReportZpz2025DataDto[] data, int startRowIndex, int endRowIndex, string form)
        {
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                string rowNum = ObjWorkSheet.Cells[i, 2].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var rowData = data?.SingleOrDefault(x => x.Code == rowNum);
                    if (rowData != null)
                    {
                        if (ObjWorkSheet.Cells[i, 8].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 8] = rowData.CountSmo;
                        }
                        if (ObjWorkSheet.Cells[i, 9].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 9] = rowData.CountSmoAnother;
                        }
                        if (ObjWorkSheet.Cells[i, 10].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 10] = rowData.CountAssignment;
                        }
                    }
                }
            }
        }

        private void FillTable2(ReportZpz2025DataDto[] data, int startRowIndex, int endRowIndex)
        {
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                string rowNum = ObjWorkSheet.Cells[i, 2].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var rowData = data?.SingleOrDefault(x => x.Code == rowNum);
                    if (rowData != null)
                    {
                        if (ObjWorkSheet.Cells[i, 5].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 5] = rowData.CountSmo;
                        }

                        if (ObjWorkSheet.Cells[i, 7].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 7] = rowData.CountInsured;
                        }

                        if (ObjWorkSheet.Cells[i, 8].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 8] = rowData.CountInsuredRepresentative;
                        }

                        if (ObjWorkSheet.Cells[i, 9].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 9] = rowData.CountTfoms;
                        }

                        if (ObjWorkSheet.Cells[i, 10].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 10] = rowData.CountSmoAnother;
                        }

                        if (ObjWorkSheet.Cells[i, 11].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 11] = rowData.CountProsecutor;
                        }
                    }
                }
            }
        }
    }
}
