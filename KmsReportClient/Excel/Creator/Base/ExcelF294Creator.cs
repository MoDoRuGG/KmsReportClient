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
    class ExcelF294Creator : ExcelBaseCreator<Report294>
    {
        private static readonly List<ReportDictionary> F294Dictionaries = new List<ReportDictionary> {
            new ReportDictionary {
                TableName = "Таблица 1",
                StartRow = 10,
                EndRow = 32,
                RowNumIndex = 2,
                ColumnStartIndex = 4,
                Index = 2
            },
            new ReportDictionary {
                TableName = "Таблица 2",
                StartRow = 8,
                EndRow = 27,
                RowNumIndex = 2,
                ColumnStartIndex = 4,
                Index = 2
            },
            new ReportDictionary {
                TableName = "Таблица 3",
                StartRow = 8,
                EndRow = 31,
                RowNumIndex = 2,
                ColumnStartIndex = 6,
                Index = 16
            },
            new ReportDictionary {
                TableName = "Таблица 4",
                StartRow = 8,
                EndRow = 23,
                RowNumIndex = 2,
                ColumnStartIndex = 6,
                Index = 16
            },
            new ReportDictionary {
                TableName = "Таблица 5",
                StartRow = 8,
                EndRow = 18,
                RowNumIndex = 2,
                ColumnStartIndex = 6,
                Index = 16
            },
            new ReportDictionary {
                TableName = "Таблица 6",
                StartRow = 8,
                EndRow = 18,
                RowNumIndex = 2,
                ColumnStartIndex = 6,
                Index = 13
            },
            new ReportDictionary {
                TableName = "Таблица 7",
                StartRow = 8,
                EndRow = 31,
                RowNumIndex = 2,
                ColumnStartIndex = 4,
                Index = 2
            },
            new ReportDictionary {
                TableName = "Таблица 8",
                StartRow = 6,
                EndRow = 13,
                RowNumIndex = 2,
                ColumnStartIndex = 4,
                Index = 1
            },
            new ReportDictionary {
                TableName = "Таблица 9",
                StartRow = 5,
                EndRow = 15,
                RowNumIndex = 2,
                ColumnStartIndex = 4,
                Index = 2
            },
            new ReportDictionary {
                TableName = "Эффективность",
                StartRow = 5,
                EndRow = 42,
                RowNumIndex = 3,
                ColumnStartIndex = 4,
                Index = 2
            }
        };

        public ExcelF294Creator(string filename, ExcelForm reportName, string header, string filialName) :
            base(filename, reportName, header, filialName, false)
        {
        }

        protected override void FillReport(Report294 report, Report294 yearReport)
        {
            string reportMonths = YymmUtils.GetMonth(report.Yymm.Substring(2, 2));
            string reportYear = report.Yymm.Substring(0, 2);

            ObjWorkSheet.Cells[3, 1] = $"за {reportMonths} 20{reportYear} года";
            ObjWorkSheet.Cells[4, 1] = FilialName;

            FillTables(report, yearReport, 0, ObjWorkSheet, ObjWorkBook);

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[9];
            FinishFilling(ObjWorkSheet);
        }


        internal void FillTables(Report294 report, Report294 yearReport, int month, Worksheet workSheet,
            Workbook workBook)
        {
            int i = 1;
            foreach (var themeData in report.ReportDataList.OrderBy(x => x.Theme))
            {
                var yearTheme = yearReport?.ReportDataList
                    .Where(x => x.Theme == themeData.Theme)
                    .SelectMany(x => x.Data)
                    .ToArray();
                var monthData = themeData.Data;
                var dict = F294Dictionaries.Single(x => x.TableName == themeData.Theme);
                int startColumn = dict.ColumnStartIndex + month * dict.Index;

                switch (themeData.Theme)
                {
                    case "Таблица 1":
                    case "Таблица 2":
                    case "Таблица 7":
                    case "Таблица 9":                  
                        FillTable1279(monthData, yearTheme, dict.RowNumIndex, startColumn, dict.StartRow, dict.EndRow,
                            workSheet);
                        break;
                    case "Таблица 3":
                    case "Таблица 4":
                    case "Таблица 5":
                        FillTable345(monthData, yearTheme, dict.RowNumIndex, startColumn, dict.StartRow, dict.EndRow,
                            workSheet);
                        break;
                    case "Таблица 6":
                        FillTable6(monthData, yearTheme, dict.RowNumIndex, startColumn, dict.StartRow, dict.EndRow,
                            workSheet);
                        break;
                    case "Таблица 8":
                        FillTable8(monthData, dict.RowNumIndex, startColumn, dict.StartRow, dict.EndRow, workSheet);
                        break;
                }

                if (i < 9)
                {
                    workSheet = (Worksheet)workBook.Sheets[++i];
                }
            }
        }

        private void FillTable1279(Report294DataDto[] report,
            Report294DataDto[] yearReport,
            int rowNumColumnIndex,
            int columnStartIndex,
            int startRow,
            int endPosition,
            Worksheet workSheet)
        {
            for (int i = startRow; i <= endPosition; i++)
            {
                string rowNum = workSheet.Cells[i, rowNumColumnIndex].Text;
                if (string.IsNullOrEmpty(rowNum))
                {
                    continue;
                }

                var dataM = report?.SingleOrDefault(x => x.RowNum == rowNum);
                var dataY = yearReport?.SingleOrDefault(x => x.RowNum == rowNum);
                if (dataM == null && dataY == null)
                    continue;

                workSheet.Cells[i, columnStartIndex] = dataM?.CountPpl ?? 0;
                workSheet.Cells[i, columnStartIndex + 1] = dataY?.CountPpl ?? 0;
            }
        }

        private void FillTable8(Report294DataDto[] report,
            int rowNumColumnIndex,
            int columnStartIndex,
            int startRow,
            int endRow,
            Worksheet workSheet)
        {
            for (int i = startRow; i <= endRow; i++)
            {
                string rowNum = workSheet.Cells[i, rowNumColumnIndex].Text;
                if (string.IsNullOrEmpty(rowNum))
                {
                    continue;
                }

                var dataM = report?.SingleOrDefault(x => x.RowNum == rowNum);
                if (dataM == null)
                    continue;
                workSheet.Cells[i, columnStartIndex] = dataM.CountPpl;
            }
        }

        private void FillTable345(Report294DataDto[] report,
            Report294DataDto[] yearReport,
            int rowNumColumnIndex,
            int columnStartIndex,
            int startRow,
            int endRow,
            Worksheet workSheet)
        {
            for (int i = startRow; i <= endRow; i++)
            {
                string rowNum = workSheet.Cells[i, rowNumColumnIndex].Text;
                if (string.IsNullOrEmpty(rowNum))
                {
                    continue;
                }

                var dataM = report?.SingleOrDefault(x => x.RowNum == rowNum);
                var dataY = yearReport?.SingleOrDefault(x => x.RowNum == rowNum);
                if (dataM == null && dataY == null)
                    continue;

                workSheet.Cells[i, columnStartIndex] = dataM?.CountSms ?? 0;
                workSheet.Cells[i, columnStartIndex + 2] = dataM?.CountPost ?? 0;
                workSheet.Cells[i, columnStartIndex + 4] = dataM?.CountPhone ?? 0;
                workSheet.Cells[i, columnStartIndex + 6] = dataM?.CountMessengers ?? 0;
                workSheet.Cells[i, columnStartIndex + 8] = dataM?.CountEmail ?? 0;
                workSheet.Cells[i, columnStartIndex + 10] = dataM?.CountAddress ?? 0;
                workSheet.Cells[i, columnStartIndex + 12] = dataM?.CountAnother ?? 0;

                workSheet.Cells[i, columnStartIndex + 1] = dataY?.CountSms ?? 0;
                workSheet.Cells[i, columnStartIndex + 3] = dataY?.CountPost ?? 0;
                workSheet.Cells[i, columnStartIndex + 5] = dataY?.CountPhone ?? 0;
                workSheet.Cells[i, columnStartIndex + 7] = dataY?.CountMessengers ?? 0;
                workSheet.Cells[i, columnStartIndex + 9] = dataY?.CountEmail ?? 0;
                workSheet.Cells[i, columnStartIndex + 11] = dataY?.CountAddress ?? 0;
                workSheet.Cells[i, columnStartIndex + 13] = dataY?.CountAnother ?? 0;
            }
        }

        private void FillTable6(Report294DataDto[] report,
            Report294DataDto[] yearReport,
            int rowNumColumnIndex,
            int columnStartIndex,
            int startRow,
            int endRow,
            Worksheet workSheet)
        {
            for (int i = startRow; i <= endRow; i++)
            {
                string rowNum = workSheet.Cells[i, rowNumColumnIndex].Text;
                if (string.IsNullOrEmpty(rowNum))
                {
                    continue;
                }

                var dataM = report?.SingleOrDefault(x => x.RowNum == rowNum);
                var dataY = yearReport?.SingleOrDefault(x => x.RowNum == rowNum);
                if (dataM == null && dataY == null)
                    continue;

                workSheet.Cells[i, columnStartIndex] = dataM?.CountOncologicalDisease ?? 0;
                workSheet.Cells[i, columnStartIndex + 2] = dataM?.CountEndocrineDisease ?? 0;
                workSheet.Cells[i, columnStartIndex + 4] = dataM?.CountBronchoDisease ?? 0;
                workSheet.Cells[i, columnStartIndex + 6] = dataM?.CountBloodDisease ?? 0;
                workSheet.Cells[i, columnStartIndex + 8] = dataM?.CountAnotherDisease ?? 0;

                workSheet.Cells[i, columnStartIndex + 1] = dataY?.CountOncologicalDisease ?? 0;
                workSheet.Cells[i, columnStartIndex + 3] = dataY?.CountEndocrineDisease ?? 0;
                workSheet.Cells[i, columnStartIndex + 5] = dataY?.CountBronchoDisease ?? 0;
                workSheet.Cells[i, columnStartIndex + 7] = dataY?.CountBloodDisease ?? 0;
                workSheet.Cells[i, columnStartIndex + 9] = dataY?.CountAnotherDisease ?? 0;
            }
        }

        private void FinishFilling(Worksheet workSheet)
        {
            workSheet.Cells[18, 3] = CurrentUser.Director;
            workSheet.Cells[21, 1] = "Дата: " + DateTime.Today.ToShortDateString();
            if (!string.IsNullOrEmpty(CurrentUser.DirectorPhone))
            {
                var code = GetPhoneCode(CurrentUser.DirectorPhone);
                var number = GetPhoneNumber(CurrentUser.DirectorPhone);
                workSheet.Cells[21, 4] = $"+7 ({code}) {number}";
            }

            workSheet.Cells[24, 3] = CurrentUser.UserName;
            workSheet.Cells[27, 1] = CurrentUser.Email ?? "";
            if (!string.IsNullOrEmpty(CurrentUser.Phone))
            {
                var code = GetPhoneCode(CurrentUser.Phone);
                var number = GetPhoneNumber(CurrentUser.Phone);
                workSheet.Cells[27, 4] = $"+7 ({code}) {number}";
            }
        }
    }
}