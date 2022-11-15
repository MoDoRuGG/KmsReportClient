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
    class ExcelPgQCreator : ExcelBaseCreator<ReportPg>
    {
        private readonly List<ReportDictionary> _pgDictionaries = new List<ReportDictionary> {
            new ReportDictionary {TableName = "Таблица 1", StartRow = 12, EndRow = 70, Index = 1},
            new ReportDictionary {TableName = "Таблица 2", StartRow = 7, EndRow = 28, Index = 2},
            new ReportDictionary {TableName = "Таблица 3", StartRow = 7, EndRow = 35, Index = 3},
            new ReportDictionary {TableName = "Таблица 4", StartRow = 7, EndRow = 11, Index = 4},
            new ReportDictionary {TableName = "Таблица 5", StartRow = 7, EndRow = 34, Index = 5},
            new ReportDictionary {TableName = "Таблица 6", StartRow = 7, EndRow = 49, Index = 6},
            new ReportDictionary {TableName = "Таблица 8", StartRow = 7, EndRow = 72, Index = 7},
            new ReportDictionary {TableName = "Таблица 10", StartRow = 7, EndRow = 49, Index = 8},
            new ReportDictionary {TableName = "Таблица 11", StartRow = 7, EndRow = 27, Index = 9},
            new ReportDictionary {TableName = "Таблица 12", StartRow = 7, EndRow = 24, Index = 10},
            new ReportDictionary {TableName = "Таблица 13", StartRow = 7, EndRow = 21, Index = 11},
            new ReportDictionary {TableName = "Таблица 1Л", StartRow = 5, EndRow = 28, Index = 12},
            new ReportDictionary {TableName = "Таблица 2Л", StartRow = 5, EndRow = 30, Index = 13},
        };

        public ExcelPgQCreator(
            string filename,
            ExcelForm reportName,
            string header,
            string filialName) : base(filename, reportName, header, filialName, false) { }

        protected override void FillReport(ReportPg report, ReportPg yearReport)
        {
            string reportMonths = YymmUtils.GetMonth(report.Yymm.Substring(2, 2));
            string reportYear = report.Yymm.Substring(0, 2);

            ObjWorkSheet.Cells[3, 1] = $"за {reportMonths} 20{reportYear} года";
            ObjWorkSheet.Cells[4, 1] = FilialName;

            foreach (var themeData in report.ReportDataList.OrderBy(x => x.Theme))
            {
                var dict = _pgDictionaries.FirstOrDefault(x => x.TableName == themeData.Theme);
                if (dict == null)
                {
                    continue;
                }
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[dict.Index];
                var data = themeData.Data;
                switch (themeData.Theme)
                {
                    case "Таблица 1":
                    case "Таблица 11":
                    case "Таблица 12":
                        FillTable1(data, dict.StartRow, dict.EndRow, themeData.Theme);
                        break;
                    case "Таблица 2":
                    case "Таблица 3":
                        FillTable2(data, dict.StartRow, dict.EndRow);
                        break;
                    case "Таблица 4":
                    case "Таблица 10":
                    case "Таблица 13":
                        FillTable4(data, dict.StartRow, dict.EndRow, themeData.Theme);
                        break;
                    case "Таблица 6":
                    case "Таблица 8":
                        FillTable6(data, dict.StartRow, dict.EndRow);
                        break;
                    case "Таблица 5":
                        FillTable5(data, dict.StartRow, dict.EndRow);
                        break;
                    case "Таблица 1Л":
                    case "Таблица 2Л":
                        FillTableLetal(data, dict.StartRow, dict.EndRow);
                            break;
                }
            }

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[13];
            FinishPg();
        }

        private void FillTable1(ReportPgDataDto[] data, int startRowIndex, int endRowIndex, string theme)
        {
            int firstColumnIndex;
            int seconfColumnIndex;
            switch (theme)
            {
                case "Таблица 1":
                    firstColumnIndex = 8;
                    seconfColumnIndex = 9;
                    break;
                case "Таблица 11":
                    firstColumnIndex = 7;
                    seconfColumnIndex = 9;
                    break;
                default:
                    firstColumnIndex = 5;
                    seconfColumnIndex = 6;
                    break;
            }

            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                string rowNum = ObjWorkSheet.Cells[i, 2].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var rowData = data?.SingleOrDefault(x => x.Code == rowNum);
                    if (rowData != null)
                    {
                        ObjWorkSheet.Cells[i, firstColumnIndex] = rowData.CountSmo;
                        ObjWorkSheet.Cells[i, seconfColumnIndex] = rowData.CountSmoAnother;
                    }
                }
            }
        }

        private void FillTable2(ReportPgDataDto[] data, int startRowIndex, int endRowIndex)
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

        private void FillTable4(ReportPgDataDto[] data, int startRowIndex, int endRowIndex, string form)
        {
            int columnIndex = form == "Таблица 13" ? 4 : 5;
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

        private void FillTable6(ReportPgDataDto[] data, int startRowIndex, int endRowIndex)
        {
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                string rowNum = ObjWorkSheet.Cells[i, 2].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var rowData = data?.SingleOrDefault(x => x.Code == rowNum);
                    if (rowData != null)
                    {
                        if (ObjWorkSheet.Cells[i, 4].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 4] = rowData.CountOutOfSmo;
                        }

                        if (ObjWorkSheet.Cells[i, 5].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 5] = rowData.CountAmbulatory;
                        }

                        if (ObjWorkSheet.Cells[i, 6].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 6] = rowData.CountDs;
                        }

                        if (ObjWorkSheet.Cells[i, 7].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 7] = rowData.CountDsVmp;
                        }

                        if (ObjWorkSheet.Cells[i, 8].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 8] = rowData.CountStac;
                        }

                        if (ObjWorkSheet.Cells[i, 9].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 9] = rowData.CountStacVmp;
                        }

                        if (ObjWorkSheet.Cells[i, 11].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 11] = rowData.CountOutOfSmoAnother;
                        }

                        if (ObjWorkSheet.Cells[i, 12].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 12] = rowData.CountAmbulatoryAnother;
                        }

                        if (ObjWorkSheet.Cells[i, 13].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 13] = rowData.CountDsAnother;
                        }

                        if (ObjWorkSheet.Cells[i, 14].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 14] = rowData.CountDsVmpAnother;
                        }

                        if (ObjWorkSheet.Cells[i, 15].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 15] = rowData.CountStacAnother;
                        }

                        if (ObjWorkSheet.Cells[i, 16].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 16] = rowData.CountStacVmpAnother;
                        }
                    }
                }
            }
        }

        private void FillTable5(ReportPgDataDto[] data, int startRowIndex, int endRowIndex)
        {
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                string rowNum = ObjWorkSheet.Cells[i, 2].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var rowData = data?.SingleOrDefault(x => x.Code == rowNum);
                    if (rowData != null)
                    {
                        if (ObjWorkSheet.Cells[i, 4].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 4] = rowData.CountOutOfSmo;
                        }

                        if (ObjWorkSheet.Cells[i, 5].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 5] = rowData.CountAmbulatory;
                        }

                        if (ObjWorkSheet.Cells[i, 6].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 6] = rowData.CountDs;
                        }

                        if (ObjWorkSheet.Cells[i, 7].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 7] = rowData.CountDsVmp;
                        }

                        if (ObjWorkSheet.Cells[i, 8].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 8] = rowData.CountStac;
                        }

                        if (ObjWorkSheet.Cells[i, 9].Text != "X")
                        {
                            ObjWorkSheet.Cells[i, 9] = rowData.CountStacVmp;
                        }
                    }
                }
            }
        }

        private void FillTableLetal(ReportPgDataDto[] data, int startRowIndex, int endRowIndex)
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

        private void FinishPg()
        {
            ObjWorkSheet.Cells[24, 3] = CurrentUser.Director;
            ObjWorkSheet.Cells[27, 1] = "Дата: " + DateTime.Today.ToShortDateString();
            if (!string.IsNullOrEmpty(CurrentUser.DirectorPhone))
            {
                var code = GetPhoneCode(CurrentUser.DirectorPhone);
                var number = GetPhoneNumber(CurrentUser.DirectorPhone);
                ObjWorkSheet.Cells[27, 4] = $"+7 ({code}) {number}";
            }

            ObjWorkSheet.Cells[30, 3] = CurrentUser.UserName;
            ObjWorkSheet.Cells[33, 1] = CurrentUser.Email ?? "";
            if (!string.IsNullOrEmpty(CurrentUser.Phone))
            {
                var code = GetPhoneCode(CurrentUser.Phone);
                var number = GetPhoneNumber(CurrentUser.Phone);
                ObjWorkSheet.Cells[33, 4] = $"+7 ({code}) {number}";
            }
        }

    }
}
