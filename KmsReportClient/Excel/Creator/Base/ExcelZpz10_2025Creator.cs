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
    class ExcelZpz10_2025Creator : ExcelBaseCreator<ReportZpz2025>
    {
        private readonly List<ReportDictionary> _zpzDictionaries = new List<ReportDictionary> {
            new ReportDictionary {TableName = "Таблица 10", StartRow = 7, EndRow = 107, Index = 1}  
        };

        public ExcelZpz10_2025Creator(
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
                var dict = _zpzDictionaries.FirstOrDefault(x => x.TableName == themeData.Theme);
                if (dict == null)
                {
                    // Обработка ошибки: лист не найден
                    Console.WriteLine($"Ошибка: Словарь для темы '{themeData.Theme}' не найден.");
                    continue; // Пропуск текущей итерации
                }
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[dict.Index];
                var data = themeData.Data;
                switch (themeData.Theme)
                {
                    case "Таблица 10":
                        FillTable10(data, dict.StartRow, dict.EndRow, themeData.Theme);
                        break;
                }
            }

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            FinishZpz();
        }


        private void FillTable10(ReportZpz2025DataDto[] data, int startRowIndex, int endRowIndex, string form)
        {
            var columnIndex = form switch
            {
                "Таблица 10" => 7,
            };
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                string rowNum = ObjWorkSheet.Cells[i, 2].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var rowData = data?.SingleOrDefault(x => x.Code == rowNum);
                    if (rowData != null)
                    {
                        ObjWorkSheet.Cells[i, columnIndex] = rowData.CountSmoAnother;
                        if (form == "Таблица 10")
                        ObjWorkSheet.Cells[i, columnIndex+1] = rowData.CountSmo;
                    }
                }
            }
        }



        private void FinishZpz()
        {
            ObjWorkSheet.Cells[110, 3] = CurrentUser.Director;
            ObjWorkSheet.Cells[113, 1] = "Дата: " + DateTime.Today.ToShortDateString();
            if (!string.IsNullOrEmpty(CurrentUser.DirectorPhone))
            {
                var code = GetPhoneCode(CurrentUser.DirectorPhone);
                var number = GetPhoneNumber(CurrentUser.DirectorPhone);
                ObjWorkSheet.Cells[113, 4] = $"+7 ({code}) {number}";
            }

            ObjWorkSheet.Cells[116, 3] = CurrentUser.UserName;
            ObjWorkSheet.Cells[119, 1] = CurrentUser.Email ?? "";
            if (!string.IsNullOrEmpty(CurrentUser.Phone))
            {
                var code = GetPhoneCode(CurrentUser.Phone);
                var number = GetPhoneNumber(CurrentUser.Phone);
                ObjWorkSheet.Cells[119, 4] = $"+7 ({code}) {number}";
            }
        }

    }
}
