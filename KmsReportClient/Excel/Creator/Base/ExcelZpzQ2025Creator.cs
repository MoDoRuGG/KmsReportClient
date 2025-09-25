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
    class ExcelZpzQ2025Creator : ExcelBaseCreator<ReportZpz2025>
    {
        private readonly List<ReportDictionary> _zpzDictionaries = new List<ReportDictionary> {
            new ReportDictionary {TableName = "Результаты МЭК", StartRow = 6, EndRow = 6, Index = 1},
            new ReportDictionary {TableName = "Таблица 6", StartRow = 7, EndRow = 187, Index = 2},
            new ReportDictionary {TableName = "Таблица 7", StartRow = 7, EndRow = 407, Index = 3},
            new ReportDictionary {TableName = "Таблица 8", StartRow = 6, EndRow = 484, Index = 4},
            new ReportDictionary {TableName = "Таблица 9", StartRow = 6, EndRow = 38, Index = 5},
            new ReportDictionary {TableName = "Оплата МП", StartRow = 6, EndRow = 6, Index = 6},
            new ReportDictionary {TableName = "Таблица 1Л", StartRow = 5, EndRow = 28, Index = 7},
            new ReportDictionary {TableName = "Таблица 2Л", StartRow = 5, EndRow = 30, Index = 8},
        };

        public ExcelZpzQ2025Creator(
            string filename,
            ExcelForm reportName,
            string header,
            string filialName) : base(filename, reportName, header, filialName, false) { }

        protected override void FillReport(ReportZpz2025 report, ReportZpz2025 yearReport)
        {
            string reportMonths = YymmUtils.GetMonth(report.Yymm.Substring(2, 2));
            string reportYear = report.Yymm.Substring(0, 2);

            //ObjWorkSheet.Cells[3, 1] = $"за {reportMonths} 20{reportYear} года";
            //ObjWorkSheet.Cells[4, 1] = FilialName;

            foreach (var themeData in report.ReportDataList.OrderBy(x => x.Theme))
            {
                //var dict = _zpzDictionaries.FirstOrDefault(x => x.TableName == themeData.Theme);
                var dict = _zpzDictionaries.Single(x => x.TableName == themeData.Theme);
                //if (dict == null)
                //{
                //    continue;
                //}
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[dict.Index];
                var data = themeData.Data;
                switch (themeData.Theme)
                {
                    case "Таблица 9":
                        FillTable9(data, dict.StartRow, dict.EndRow, themeData.Theme);
                        break;
                    case "Таблица 8":
                        FillTable8(data, dict.StartRow, dict.EndRow, themeData.Theme);
                        break;
                    case "Таблица 6":
                    case "Таблица 7":
                        FillTable67(data, dict.StartRow, dict.EndRow, themeData.Theme);
                        break;
                    case "Результаты МЭК":
                    case "Оплата МП":
                        FillTable5A8A(data, dict.StartRow, dict.EndRow, themeData.Theme);
                        break;
                    case "Таблица 1Л":
                    case "Таблица 2Л":
                        FillTableLetal(data, dict.StartRow, dict.EndRow, themeData.Theme);
                            break;
                }
            }

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[5];
            FinishZpz();
        }

        private void FillTable9(ReportZpz2025DataDto[] data, int startRowIndex, int endRowIndex, string theme)
        {
            int firstColumnIndex;
            int seconfColumnIndex;
            firstColumnIndex = 7;
            seconfColumnIndex = 9;

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

        //private void FillTable2(ReportZpzDataDto[] data, int startRowIndex, int endRowIndex)
        //{
        //    for (int i = startRowIndex; i <= endRowIndex; i++)
        //    {
        //        string rowNum = ObjWorkSheet.Cells[i, 2].Text;
        //        if (!string.IsNullOrEmpty(rowNum))
        //        {
        //            var rowData = data?.SingleOrDefault(x => x.Code == rowNum);
        //            if (rowData != null)
        //            {
        //                if (ObjWorkSheet.Cells[i, 5].Text != "x")
        //                {
        //                    ObjWorkSheet.Cells[i, 5] = rowData.CountSmo;
        //                }

        //                if (ObjWorkSheet.Cells[i, 7].Text != "x")
        //                {
        //                    ObjWorkSheet.Cells[i, 7] = rowData.CountInsured;
        //                }

        //                if (ObjWorkSheet.Cells[i, 8].Text != "x")
        //                {
        //                    ObjWorkSheet.Cells[i, 8] = rowData.CountInsuredRepresentative;
        //                }

        //                if (ObjWorkSheet.Cells[i, 9].Text != "x")
        //                {
        //                    ObjWorkSheet.Cells[i, 9] = rowData.CountTfoms;
        //                }

        //                if (ObjWorkSheet.Cells[i, 10].Text != "x")
        //                {
        //                    ObjWorkSheet.Cells[i, 10] = rowData.CountSmoAnother;
        //                }

        //                if (ObjWorkSheet.Cells[i, 11].Text != "x")
        //                {
        //                    ObjWorkSheet.Cells[i, 11] = rowData.CountProsecutor;
        //                }
        //            }
        //        }
        //    }
        //}

        private void FillTable8(ReportZpz2025DataDto[] data, int startRowIndex, int endRowIndex, string form)
        {
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                string rowNum = ObjWorkSheet.Cells[i, 2].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var rowData = data?.SingleOrDefault(x => x.Code == rowNum);
                    if (rowData != null)
                    {

                        if (ObjWorkSheet.Cells[i, 11].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 11] = rowData.CountOutOfSmoAnother;
                        }

                        if (ObjWorkSheet.Cells[i, 12].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 12] = rowData.CountAmbulatoryAnother;
                        }

                        if (ObjWorkSheet.Cells[i, 13].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 13] = rowData.CountDsAnother;
                        }

                        if (ObjWorkSheet.Cells[i, 14].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 14] = rowData.CountDsVmpAnother;
                        }

                        if (ObjWorkSheet.Cells[i, 15].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 15] = rowData.CountStacAnother;
                        }

                        if (ObjWorkSheet.Cells[i, 16].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 16] = rowData.CountStacVmpAnother;
                        }
                    }
                }
            }
        }

        private void FillTable67(ReportZpz2025DataDto[] data, int startRowIndex, int endRowIndex, string form)
        {
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                string rowNum = ObjWorkSheet.Cells[i, 2].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var rowData = data?.SingleOrDefault(x => x.Code == rowNum);
                    if (rowData != null)
                    {
                        if (ObjWorkSheet.Cells[i, 4].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 4] = rowData.CountOutOfSmo;
                        }

                        if (ObjWorkSheet.Cells[i, 5].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 5] = rowData.CountAmbulatory;
                        }

                        if (ObjWorkSheet.Cells[i, 6].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 6] = rowData.CountDs;
                        }

                        if (ObjWorkSheet.Cells[i, 7].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 7] = rowData.CountDsVmp;
                        }

                        if (ObjWorkSheet.Cells[i, 8].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 8] = rowData.CountStac;
                        }

                        if (ObjWorkSheet.Cells[i, 9].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 9] = rowData.CountStacVmp;
                        }

                        if (ObjWorkSheet.Cells[i, 11].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 11] = rowData.CountOutOfSmoAnother;
                        }

                        if (ObjWorkSheet.Cells[i, 12].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 12] = rowData.CountAmbulatoryAnother;
                        }

                        if (ObjWorkSheet.Cells[i, 13].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 13] = rowData.CountDsAnother;
                        }

                        if (ObjWorkSheet.Cells[i, 14].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 14] = rowData.CountDsVmpAnother;
                        }

                        if (ObjWorkSheet.Cells[i, 15].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 15] = rowData.CountStacAnother;
                        }

                        if (ObjWorkSheet.Cells[i, 16].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 16] = rowData.CountStacVmpAnother;
                        }
                    }
                }
            }
        }

        private void FillTable5A8A(ReportZpz2025DataDto[] data, int startRowIndex, int endRowIndex, string form)
        {
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                if (data != null)
                {

                    if (ObjWorkSheet.Cells[i, 2].Text != "x")
                    {
                        ObjWorkSheet.Cells[i, 2] = data[0].Code;
                    }
                    if (ObjWorkSheet.Cells[i, 3].Text != "x")
                    {
                        ObjWorkSheet.Cells[i, 3] = data[0].CountSmo;
                    }
                    if (ObjWorkSheet.Cells[i, 4].Text != "x")
                    {
                        ObjWorkSheet.Cells[i, 4] = data[0].CountSmoAnother;
                    }

                    if (ObjWorkSheet.Cells[i, 5].Text != "x")
                    {
                        ObjWorkSheet.Cells[i, 5] = data[0].CountInsured;
                    }

                    if (ObjWorkSheet.Cells[i, 6].Text != "x")
                    {
                        ObjWorkSheet.Cells[i, 6] = data[0].CountInsuredRepresentative;
                    }

                    if (ObjWorkSheet.Cells[i, 7].Text != "x")
                    {
                        ObjWorkSheet.Cells[i, 7] = data[0].CountTfoms;
                    }

                    if (ObjWorkSheet.Cells[i, 8].Text != "x")
                    {
                        ObjWorkSheet.Cells[i, 8] = data[0].CountProsecutor;
                    }

                    if (ObjWorkSheet.Cells[i, 9].Text != "x")
                    {
                        ObjWorkSheet.Cells[i, 9] = data[0].CountOutOfSmo;
                    }
                }
            }
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
                        if (ObjWorkSheet.Cells[i, 8].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 8] = rowData.CountAmbulatory;
                        }

                        if (ObjWorkSheet.Cells[i, 9].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 9] = rowData.CountStac;
                        }

                        if (ObjWorkSheet.Cells[i, 10].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 10] = rowData.CountDs;
                        }

                        if (ObjWorkSheet.Cells[i, 11].Text != "x")
                        {
                            ObjWorkSheet.Cells[i, 11] = rowData.CountOutOfSmoAnother;
                        }

                        if (ObjWorkSheet.Cells[i, 12].Text != "x")
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
