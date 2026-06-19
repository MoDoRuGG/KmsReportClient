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
    class ExcelDispReprodHealthCreator : ExcelBaseCreator<ReportDispReprodHealth>
    {
        private readonly List<ReportDictionary> _Dictionaries = new List<ReportDictionary> {
            new ReportDictionary {TableName = "Дисп РЗ", StartRow = 9, EndRow = 27, Index = 1}
        };

        public ExcelDispReprodHealthCreator(
            string filename,
            ExcelForm reportName,
            string header,
            string filialName) : base(filename, reportName, header, filialName, false) { }

        protected override void FillReport(ReportDispReprodHealth report, ReportDispReprodHealth yearReport)
        {
            string reportMonths = YymmUtils.GetMonth(report.Yymm.Substring(2, 2));
            string reportYear = report.Yymm.Substring(0, 2);

            //ObjWorkSheet.Cells[3, 1] = $"за {reportMonths} 20{reportYear} года";
            //ObjWorkSheet.Cells[4, 1] = FilialName;

            foreach (var themeData in report.ReportDataList.OrderBy(x => x.Theme))
            {
                var dict = _Dictionaries.FirstOrDefault(x => x.TableName == themeData.Theme);
                if (dict == null)
                {
                    // Обработка ошибки: лист не найден
                    Console.WriteLine($"Ошибка: Словарь для темы '{themeData.Theme}' не найден.");
                    continue; // Пропуск текущей итерации
                }
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[dict.Index];
                var data = themeData.Data;

                FillTable(data, dict.StartRow, dict.EndRow, themeData.Theme);
                break;

            }

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
        }


        private void FillTable(ReportDispReprodHealthDataDto[] data, int startRowIndex, int endRowIndex, string form)
        {
            var columnIndex = 7;

            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                string rowNum = ObjWorkSheet.Cells[i, 2].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var rowData = data?.SingleOrDefault(x => x.Code == rowNum);
                    if (rowData != null)
                    {
                        ObjWorkSheet.Cells[i, columnIndex] = rowData.YearlySum;
                        ObjWorkSheet.Cells[i, columnIndex + 1] = rowData.ForPeriod;
                    }
                }
            }
        }
    }
}
