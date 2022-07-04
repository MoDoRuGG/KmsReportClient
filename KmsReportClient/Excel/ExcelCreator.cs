using System;
using System.Collections.Generic;
using System.Linq;
using KmsReportClient.Report;
using KmsReportClient.Utils;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel
{
    class ExcelCreator<T>
    {
        private readonly string filename;
        private readonly int startPosition;
        private readonly string reportName;
        private readonly string header;

        private Application objExcel;
        private Workbook objWorkBook;
        private Worksheet objWorkSheet;

        public ExcelCreator(string filename, int startPosition, string reportName, string header)
        {
            this.filename = filename;
            this.startPosition = startPosition;
            this.reportName = reportName;
            this.header = header;
        }

        public void CreateReport(T report, T yearReport)
        {
            objExcel = new Application
            {
                DisplayAlerts = false
            };
            objWorkBook = objExcel.Workbooks.Open(reportName);
            objWorkSheet = (Worksheet)objWorkBook.Sheets[1];

            try
            {
                if (reportName == TemplateName.Iizl)
                {
                    FillIizl(report as ReportIizl);
                }
                else if (reportName == TemplateName.F262)
                {
                    Fill262(report as Report262, yearReport as Report262);
                }
                else if (reportName == TemplateName.F294)
                {
                    FillFilial294(report as Report294, yearReport as Report294);
                }
                else if (reportName == TemplateName.Pg)
                {
                    FillFilialPg(report as ReportPg);
                }
                else if (reportName == TemplateName.Cons262T1)
                {
                    FillConsolidateReport262T1(report as External.CReport262Table1[]);
                }
                else if (reportName == TemplateName.Cons262T2)
                {
                    FillConsolidateReport262T2(report as External.CReport262Table2[],
                        yearReport as External.CReport262Table2[]);
                }
                else if (reportName == TemplateName.Cons262T3)
                {
                    FillConsolidateReport262T3(report as External.CReport262Table3[]);
                }
                else if (reportName == TemplateName.Cons294)
                {
                    FillConsolidateReport294(report as List<Report294>);
                }
                else if (reportName == TemplateName.ControlZpz)
                {
                    FillControlZpz(report as External.CReportPg[]);
                }
                objWorkBook.SaveAs(filename, Type.Missing, Type.Missing, Type.Missing,
                                             Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange,
                                             Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            finally
            {
                objExcel.Quit();
                objWorkBook = null;
                objWorkSheet = null;
                objExcel = null;
                GC.Collect();
            }
        }

        private void FillControlZpz(External.CReportPg[] reportList)
        {
            int countReport = reportList.Length;
            int currentIndex = startPosition;
            CopyNullCells(objWorkSheet, countReport, startPosition);

            foreach (var data in reportList)
            {
                objWorkSheet.Cells[currentIndex, 1] = data.Filial;
                objWorkSheet.Cells[currentIndex, 2] = data.Bills;
                objWorkSheet.Cells[currentIndex, 3] = data.BillsOnco;
                objWorkSheet.Cells[currentIndex, 4] = data.BillsVioletion;
                objWorkSheet.Cells[currentIndex, 5] = data.PaymentBills;
                objWorkSheet.Cells[currentIndex, 6] = data.PaymentBillsOnco;
                objWorkSheet.Cells[currentIndex, 7] = data.MeeTarget;
                objWorkSheet.Cells[currentIndex, 8] = data.MeePlan;
                objWorkSheet.Cells[currentIndex, 10] = data.CaseMeeTarget;
                objWorkSheet.Cells[currentIndex, 11] = data.CaseMeePlan;
                objWorkSheet.Cells[currentIndex, 13] = data.DefectMeeTarget;
                objWorkSheet.Cells[currentIndex, 14] = data.DefectMeePlan;
                objWorkSheet.Cells[currentIndex, 16] = data.EkmpTarget;
                objWorkSheet.Cells[currentIndex, 17] = data.EkmpPlan;
                objWorkSheet.Cells[currentIndex, 19] = data.ThemeCaseEkmpPlan;
                objWorkSheet.Cells[currentIndex, 22] = data.CaseEkmpTarget;
                objWorkSheet.Cells[currentIndex, 23] = data.CaseEkmpPlan;
                objWorkSheet.Cells[currentIndex, 25] = data.DefectEkmpTarget;
                objWorkSheet.Cells[currentIndex++, 26] = data.DefectEkmpPlan;
            }
        }


        private void FillConsolidateReport294(List<Report294> reportList)
        {
            int month = 0;
            foreach (var report in reportList)
            {
                objWorkSheet = (Worksheet)objWorkBook.Sheets[1];
                Fill294(report, null, month);
                month++;
            }
        }

        private void FillFilial294(Report294 monthReport, Report294 yearReport)
        {
            string reportMonths = GlobalUtils.GetMonth(monthReport.Yymm.Substring(2, 2));
            string reportYear = monthReport.Yymm.Substring(0, 2);

            objWorkSheet.Cells[3, 1] = $"за {reportMonths} 20{reportYear} года";
            objWorkSheet.Cells[4, 1] = monthReport.FilialName;

            Fill294(monthReport, yearReport, 0);

            objWorkSheet = (Worksheet)objWorkBook.Sheets[9];
            Finish294();
        }

        private void FillFilialPg(ReportPg report)
        {
            string reportMonths = GlobalUtils.GetMonth(report.Yymm.Substring(2, 2));
            string reportYear = report.Yymm.Substring(0, 2);

            objWorkSheet.Cells[3, 1] = $"за {reportMonths} 20{reportYear} года";
            objWorkSheet.Cells[4, 1] = report.FilialName;

            foreach (var themeData in report.ReportDataList.OrderBy(x => x.Theme))
            {
                var dict = pgDictionaries.Single(x => x.TableName == themeData.Theme);
                objWorkSheet = (Worksheet)objWorkBook.Sheets[dict.Index];
                var data = themeData.Data;
                switch (themeData.Theme)
                {
                    case "Таблица 1":
                    case "Таблица 11":
                    case "Таблица 12":
                        FillReportPgTable1(data,                            
                            dict.StartRowIndex,
                            dict.EndRowIndex, 
                            themeData.Theme);
                        break;
                    case "Таблица 2":
                    case "Таблица 3":
                        FillReportPgTable2(data,   
                            dict.StartRowIndex,
                            dict.EndRowIndex);
                        break;
                    case "Таблица 4":
                    case "Таблица 10":
                    case "Таблица 13":
                        FillReportPgTable4(data,
                            dict.StartRowIndex,
                            dict.EndRowIndex,
                            themeData.Theme);
                        break;
                    case "Таблица 6":
                    case "Таблица 8":
                        FillReportPgTable6(data,
                            dict.StartRowIndex,
                            dict.EndRowIndex);
                        break;
                    case "Таблица 5":
                        FillReportPgTable5(data,
                            dict.StartRowIndex,
                            dict.EndRowIndex);
                        break;
                }
            }

            objWorkSheet = (Worksheet)objWorkBook.Sheets[13];
            FinishPg();
        }

        private void FillReportPgTable1(List<ReportPgDataDto> data, int startRowIndex, int endRowIndex, string theme)
        {
            int firstColumnIndex;
            int seconfColumnIndex;
            if (theme == "Таблица 1")
            {
                firstColumnIndex = 8;
                seconfColumnIndex = 9;
            }
            else if (theme == "Таблица 11")
            {
                firstColumnIndex = 7;
                seconfColumnIndex = 9;
            }
            else
            {
                firstColumnIndex = 5;
                seconfColumnIndex = 6;
            }
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                string rowNum = objWorkSheet.Cells[i, 2].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var rowData = data.SingleOrDefault(x => x.Code == rowNum);
                    if (rowData != null)
                    {
                        objWorkSheet.Cells[i, firstColumnIndex] = rowData.CountSmo;
                        objWorkSheet.Cells[i, seconfColumnIndex] = rowData.CountSmoAnother;
                    }
                }
            }
        }

        private void FillReportPgTable2(List<ReportPgDataDto> data, int startRowIndex, int endRowIndex)
        {
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                string rowNum = objWorkSheet.Cells[i, 2].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var rowData = data.SingleOrDefault(x => x.Code == rowNum);
                    if (rowData != null)
                    {
                        if (objWorkSheet.Cells[i, 5].Text != "X") objWorkSheet.Cells[i, 5] = rowData.CountSmo;
                        if (objWorkSheet.Cells[i, 7].Text != "X") objWorkSheet.Cells[i, 7] = rowData.CountInsured;
                        if (objWorkSheet.Cells[i, 8].Text != "X") objWorkSheet.Cells[i, 8] = rowData.CountInsuredRepresentative;
                        if (objWorkSheet.Cells[i, 9].Text != "X") objWorkSheet.Cells[i, 9] = rowData.CountTfoms;
                        if (objWorkSheet.Cells[i, 10].Text != "X") objWorkSheet.Cells[i, 10] = rowData.CountSmoAnother;
                        if (objWorkSheet.Cells[i, 11].Text != "X") objWorkSheet.Cells[i, 11] = rowData.CountProsecutor;
                    }
                }
            }
        }

        private void FillReportPgTable4(List<ReportPgDataDto> data, int startRowIndex, int endRowIndex, string form)
        {
            int columnIndex = form == "Таблица 13" ? 4 : 5;
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                string rowNum = objWorkSheet.Cells[i, 2].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var rowData = data.SingleOrDefault(x => x.Code == rowNum);
                    if (rowData != null)
                    {
                        objWorkSheet.Cells[i, columnIndex] = rowData.CountSmo;
                    }
                }
            }
        }

        private void FillReportPgTable6(List<ReportPgDataDto> data, int startRowIndex, int endRowIndex)
        {
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                string rowNum = objWorkSheet.Cells[i, 2].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var rowData = data.SingleOrDefault(x => x.Code == rowNum);
                    if (rowData != null)
                    {
                        if (objWorkSheet.Cells[i, 4].Text != "X") objWorkSheet.Cells[i, 4] = rowData.CountOutOfSmo;
                        if (objWorkSheet.Cells[i, 5].Text != "X") objWorkSheet.Cells[i, 5] = rowData.CountAmbulatory;
                        if (objWorkSheet.Cells[i, 6].Text != "X") objWorkSheet.Cells[i, 6] = rowData.CountDs;
                        if (objWorkSheet.Cells[i, 7].Text != "X") objWorkSheet.Cells[i, 7] = rowData.CountDsVmp;
                        if (objWorkSheet.Cells[i, 8].Text != "X") objWorkSheet.Cells[i, 8] = rowData.CountStac;
                        if (objWorkSheet.Cells[i, 9].Text != "X") objWorkSheet.Cells[i, 9] = rowData.CountStacVmp;
                        if (objWorkSheet.Cells[i, 11].Text != "X") objWorkSheet.Cells[i, 11] = rowData.CountOutOfSmoAnother;
                        if (objWorkSheet.Cells[i, 12].Text != "X") objWorkSheet.Cells[i, 12] = rowData.CountAmbulatoryAnother;
                        if (objWorkSheet.Cells[i, 13].Text != "X") objWorkSheet.Cells[i, 13] = rowData.CountDsAnother;
                        if (objWorkSheet.Cells[i, 14].Text != "X") objWorkSheet.Cells[i, 14] = rowData.CountDsVmpAnother;
                        if (objWorkSheet.Cells[i, 15].Text != "X") objWorkSheet.Cells[i, 15] = rowData.CountStacAnother;
                        if (objWorkSheet.Cells[i, 16].Text != "X") objWorkSheet.Cells[i, 16] = rowData.CountStacVmpAnother;
                    }
                }
            }
        }

        private void FillReportPgTable5(List<ReportPgDataDto> data, int startRowIndex, int endRowIndex)
        {
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                string rowNum = objWorkSheet.Cells[i, 2].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var rowData = data.SingleOrDefault(x => x.Code == rowNum);
                    if (rowData != null)
                    {
                        if (objWorkSheet.Cells[i, 4].Text != "X") objWorkSheet.Cells[i, 4] = rowData.CountOutOfSmo;
                        if (objWorkSheet.Cells[i, 5].Text != "X") objWorkSheet.Cells[i, 5] = rowData.CountAmbulatory;
                        if (objWorkSheet.Cells[i, 6].Text != "X") objWorkSheet.Cells[i, 6] = rowData.CountDs;
                        if (objWorkSheet.Cells[i, 7].Text != "X") objWorkSheet.Cells[i, 7] = rowData.CountDsVmp;
                        if (objWorkSheet.Cells[i, 8].Text != "X") objWorkSheet.Cells[i, 8] = rowData.CountStac;
                        if (objWorkSheet.Cells[i, 9].Text != "X") objWorkSheet.Cells[i, 9] = rowData.CountStacVmp;
                    }
                }
            }
        }

        private void FinishPg()
        {
            objWorkSheet.Cells[24, 3] = CurrentUser.Director;
            objWorkSheet.Cells[27, 1] = "Дата: " + DateTime.Today.ToShortDateString();
            if (!string.IsNullOrEmpty(CurrentUser.DirectorPhone))
            {
                var code = GetPhoneCode(CurrentUser.DirectorPhone);
                var number = GetPhoneNumber(CurrentUser.DirectorPhone);
                objWorkSheet.Cells[27, 4] = $"+7 ({code}) {number}";
            }
            objWorkSheet.Cells[30, 3] = CurrentUser.UserName;
            objWorkSheet.Cells[33, 1] = CurrentUser.Email ?? "";
            if (!string.IsNullOrEmpty(CurrentUser.Phone))
            {
                var code = GetPhoneCode(CurrentUser.Phone);
                var number = GetPhoneNumber(CurrentUser.Phone);
                objWorkSheet.Cells[33, 4] = $"+7 ({code}) {number}";
            }
        }

        private void Fill294(Report294 report, Report294 yearReport, int month)
        {
            int i = 1;
            foreach (var themeData in report.ReportDataList.OrderBy(x => x.Theme))
            {
                var yearTheme = yearReport?.ReportDataList
                    .Where(x => x.Theme == themeData.Theme)
                    .SelectMany(x => x.Data)
                    .ToList() ?? null;
                var monthData = themeData.Data;
                var dict = f294Dictionaries.Single(x => x.TableName == themeData.Theme);
                int startColumn = dict.ColumnStartIndex + month * dict.Index;
                switch (themeData.Theme)
                {
                    case "Таблица 1":
                    case "Таблица 2":
                    case "Таблица 7":
                    case "Таблица 9":
                    case "Эффективность":
                        FillReport294Table1279(monthData,
                            yearTheme,
                            dict.RowNumIndex,
                            startColumn,
                            dict.StartRowIndex,
                            dict.EndRowIndex);
                        break;
                    case "Таблица 3":
                    case "Таблица 4":
                    case "Таблица 5":
                        FillReport294Table345(monthData,
                            yearTheme,
                            dict.RowNumIndex,
                            startColumn,
                            dict.StartRowIndex,
                            dict.EndRowIndex);
                        break;
                    case "Таблица 6":
                        FillReport294Table6(monthData,
                            yearTheme,
                            dict.RowNumIndex,
                            startColumn,
                            dict.StartRowIndex,
                            dict.EndRowIndex);
                        break;
                    case "Таблица 8":
                        FillReport294Table8(monthData,
                            dict.RowNumIndex,
                            startColumn,
                            dict.StartRowIndex,
                            dict.EndRowIndex);
                        break;
                }
                if (i < 10)
                {
                    objWorkSheet = (Worksheet)objWorkBook.Sheets[++i];
                }
            }
        }

        private void FillReport294Table1279(List<Report294DataDto> report,
            List<Report294DataDto> yearReport,
            int rowNumColumnIndex,
            int columnStartIndex,
            int startPosition,
            int endPosition)
        {
            for (int i = startPosition; i <= endPosition; i++)
            {
                string rowNum = objWorkSheet.Cells[i, rowNumColumnIndex].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var dataM = report.SingleOrDefault(x => x.RowNum == rowNum);
                    if (dataM != null)
                    {
                        objWorkSheet.Cells[i, columnStartIndex] = dataM.CountPpl;
                    }
                    var dataY = yearReport?.SingleOrDefault(x => x.RowNum == rowNum);
                    if (dataY != null)
                    {
                        objWorkSheet.Cells[i, columnStartIndex + 1] = dataY?.CountPpl ?? 0;
                    }
                }
            }
        }

        private void FillReport294Table8(List<Report294DataDto> report,
            int rowNumColumnIndex,
            int columnStartIndex,
            int startPosition,
            int endPosition)
        {
            for (int i = startPosition; i <= endPosition; i++)
            {
                string rowNum = objWorkSheet.Cells[i, rowNumColumnIndex].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var dataM = report.SingleOrDefault(x => x.RowNum == rowNum);
                    if (dataM == null)
                    {
                        continue;
                    }

                    objWorkSheet.Cells[i, columnStartIndex] = dataM?.CountPpl ?? 0;
                }
            }
        }

        private void FillReport294Table345(List<Report294DataDto> report,
            List<Report294DataDto> yearReport,
            int rowNumColumnIndex,
            int columnStartIndex,
            int startPosition,
            int endPosition)
        {
            for (int i = startPosition; i <= endPosition; i++)
            {
                string rowNum = objWorkSheet.Cells[i, rowNumColumnIndex].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var dataM = report.SingleOrDefault(x => x.RowNum == rowNum);
                    if (dataM != null)
                    {
                        objWorkSheet.Cells[i, columnStartIndex] = dataM?.CountSms ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 2] = dataM?.CountPost ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 4] = dataM?.CountPhone ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 6] = dataM?.CountMessangers ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 8] = dataM?.CountEmail ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 10] = dataM?.CountAddress ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 12] = dataM?.CountAnother ?? 0;
                    }

                    var dataY = yearReport?.SingleOrDefault(x => x.RowNum == rowNum);
                    if (dataY != null)
                    {
                        objWorkSheet.Cells[i, columnStartIndex + 1] = dataY?.CountSms ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 3] = dataY?.CountPost ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 5] = dataY?.CountPhone ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 7] = dataY?.CountMessangers ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 9] = dataY?.CountEmail ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 11] = dataY?.CountAddress ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 13] = dataY?.CountAnother ?? 0;
                    }
                }
            }
        }

        private void FillReport294Table6(List<Report294DataDto> report,
            List<Report294DataDto> yearReport,
            int rowNumColumnIndex,
            int columnStartIndex,
            int startPosition,
            int endPosition)
        {
            for (int i = startPosition; i <= endPosition; i++)
            {
                string rowNum = objWorkSheet.Cells[i, rowNumColumnIndex].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var dataM = report.SingleOrDefault(x => x.RowNum == rowNum);

                    if (dataM != null)
                    {
                        objWorkSheet.Cells[i, columnStartIndex] = dataM?.CountOncologicalDisease ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 2] = dataM?.CountEndocrineDisease ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 4] = dataM?.CountBronchoDisease ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 6] = dataM?.CountBloodDisease ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 8] = dataM?.CountAnotherDisease ?? 0;
                    }

                    var dataY = yearReport?.SingleOrDefault(x => x.RowNum == rowNum);
                    if (dataY != null)
                    {
                        objWorkSheet.Cells[i, columnStartIndex + 1] = dataY?.CountOncologicalDisease ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 3] = dataY?.CountEndocrineDisease ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 5] = dataY?.CountBronchoDisease ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 7] = dataY?.CountBloodDisease ?? 0;
                        objWorkSheet.Cells[i, columnStartIndex + 9] = dataY?.CountAnotherDisease ?? 0;
                    }
                }
            }
        }

        private void Finish294()
        {
            objWorkSheet.Cells[18, 3] = CurrentUser.Director;
            objWorkSheet.Cells[21, 1] = "Дата: " + DateTime.Today.ToShortDateString();
            if (!string.IsNullOrEmpty(CurrentUser.DirectorPhone))
            {
                var code = GetPhoneCode(CurrentUser.DirectorPhone);
                var number = GetPhoneNumber(CurrentUser.DirectorPhone);
                objWorkSheet.Cells[21, 4] = $"+7 ({code}) {number}";
            }
            objWorkSheet.Cells[24, 3] = CurrentUser.UserName;
            objWorkSheet.Cells[27, 1] = CurrentUser.Email ?? "";
            if (!string.IsNullOrEmpty(CurrentUser.Phone))
            {
                var code = GetPhoneCode(CurrentUser.Phone);
                var number = GetPhoneNumber(CurrentUser.Phone);
                objWorkSheet.Cells[27, 4] = $"+7 ({code}) {number}";
            }
        }

        private void Fill262(Report262 report, Report262 yearReport)
        {
            int i = 1;
            string reportMonths = "";
            string reportYear = "";
            if (report.Yymm.Length == 4)
            {
                reportMonths = GlobalUtils.GetMonth(report.Yymm.Substring(2, 2));
                reportYear = report.Yymm.Substring(0, 2);
            }
            foreach (var themeData in report.ReportDataList.OrderBy(x => x.Theme))
            {
                if (i == 1)
                {
                    int currentIndex = 20;
                    foreach (var data in themeData.Data.OrderBy(x => x.RowNum))
                    {
                        var yearTheme = yearReport.ReportDataList.Where(x => x.Theme == themeData.Theme).SelectMany(x => x.Data);
                        var yearData = yearTheme.Single(x => x.RowNum == data.RowNum);
                        objWorkSheet.Cells[currentIndex, 80] = data.CountPpl;
                        objWorkSheet.Cells[currentIndex++, 93] = yearData.CountPpl;
                    }
                    objWorkSheet.Cells[12, 19] = report.FilialName;
                    objWorkSheet.Cells[10, 40] = reportMonths;
                    objWorkSheet.Cells[10, 62] = reportYear;
                }
                else if (i == 2)
                {
                    if (themeData.Data.Count == 0)
                    {
                        continue;
                    }

                    var data = themeData.Data[0];
                    var yearTheme = yearReport.ReportDataList.Where(x => x.Theme == themeData.Theme).SelectMany(x => x.Data);
                    var yearData = yearTheme.ToArray()[0];
                    objWorkSheet.Cells[7, 56] = data.CountSms;
                    objWorkSheet.Cells[7, 65] = yearData.CountSms;
                    objWorkSheet.Cells[7, 72] = data.CountPost;
                    objWorkSheet.Cells[7, 81] = yearData.CountPost;
                    objWorkSheet.Cells[7, 88] = data.CountPhone;
                    objWorkSheet.Cells[7, 97] = yearData.CountPhone;
                    objWorkSheet.Cells[7, 104] = data.CountMessengers;
                    objWorkSheet.Cells[7, 113] = yearData.CountMessengers;
                    objWorkSheet.Cells[7, 120] = data.CountEmail;
                    objWorkSheet.Cells[7, 129] = yearData.CountEmail;
                    objWorkSheet.Cells[7, 136] = data.CountAddress;
                    objWorkSheet.Cells[7, 145] = yearData.CountAddress;
                    objWorkSheet.Cells[7, 152] = data.CountAnother;
                    objWorkSheet.Cells[7, 161] = yearData.CountAnother;

                    objWorkSheet.Cells[11, 52] = CurrentUser.Director;
                    objWorkSheet.Cells[14, 58] = GetPhoneCode(CurrentUser.DirectorPhone);
                    objWorkSheet.Cells[14, 67] = GetPhoneNumber(CurrentUser.DirectorPhone);

                    objWorkSheet.Cells[18, 52] = CurrentUser.UserName;
                    objWorkSheet.Cells[21, 58] = GetPhoneCode(CurrentUser.Phone);
                    objWorkSheet.Cells[21, 67] = GetPhoneNumber(CurrentUser.Phone);

                    objWorkSheet.Cells[21, 6] = CurrentUser.Email;

                    var date = DateTime.Today;
                    objWorkSheet.Cells[14, 8] = date.Day;
                    objWorkSheet.Cells[14, 15] = date.ToString("MMMM");
                    objWorkSheet.Cells[14, 37] = date.ToString("yy");
                }
                else
                {
                    var date = DateTime.Today;
                    objWorkSheet.Cells[8, 72] = reportMonths;
                    objWorkSheet.Cells[8, 94] = reportYear;

                    objWorkSheet.Cells[21, 52] = CurrentUser.Director;
                    objWorkSheet.Cells[24, 58] = GetPhoneCode(CurrentUser.DirectorPhone);
                    objWorkSheet.Cells[24, 67] = GetPhoneNumber(CurrentUser.DirectorPhone);

                    objWorkSheet.Cells[28, 52] = CurrentUser.UserName;
                    objWorkSheet.Cells[31, 58] = GetPhoneCode(CurrentUser.Phone);
                    objWorkSheet.Cells[31, 67] = GetPhoneNumber(CurrentUser.Phone);
                    objWorkSheet.Cells[31, 6] = CurrentUser.Email;

                    objWorkSheet.Cells[24, 8] = date.Day;
                    objWorkSheet.Cells[24, 15] = date.ToString("MMMM");
                    objWorkSheet.Cells[24, 37] = date.ToString("yy");

                    int position = 16;
                    CopyNullCells(objWorkSheet, themeData.Table3.Count, position);
                    foreach (var data in themeData.Table3)
                    {
                        objWorkSheet.Cells[position, 1] = data.Mo;
                        objWorkSheet.Cells[position, 55] = data.CountUnit;
                        objWorkSheet.Cells[position, 66] = data.CountUnitChild;
                        objWorkSheet.Cells[position, 75] = data.CountUnitWithSp;
                        objWorkSheet.Cells[position, 90] = data.CountUnitWithSpChild;
                        objWorkSheet.Cells[position, 103] = data.CountChannelSp;
                        objWorkSheet.Cells[position, 111] = data.CountChannelSpChild;
                        objWorkSheet.Cells[position, 119] = data.CountChannelPhone;
                        objWorkSheet.Cells[position, 127] = data.CountChannelPhoneChild;
                        objWorkSheet.Cells[position, 136] = data.CountChannelTerminal;
                        objWorkSheet.Cells[position, 144] = data.CountChannelTerminalChild;
                        objWorkSheet.Cells[position, 152] = data.CountChannelAnother;
                        objWorkSheet.Cells[position, 160] = data.CountChannelAnotherChild;
                        position++;
                    }
                }
                if (i < 3)
                {
                    objWorkSheet = (Worksheet)objWorkBook.Sheets[++i];
                }
            }
        }

        private void FillIizl(ReportIizl report)
        {
            foreach (var table in iizlDictionaries)
            {
                var themeData = report.ReportDataList.Single(x => x.Theme == table.TableName);
                int currentIndex = table.StartRowIndex;

                if (themeData.Theme.StartsWith("Тема"))
                {
                    string prefix = themeData.Theme.Split(' ')[1];
                    string[] suffixes = { "У", "П" };
                    foreach (var suffix in suffixes)
                    {
                        foreach (var data in themeData.Data
                            .Where(x => x.Code.StartsWith($"{prefix}-{suffix}")).OrderBy(x => x.Code))
                        {
                            objWorkSheet.Cells[currentIndex, 3] = data.CountPersFirst;
                            objWorkSheet.Cells[currentIndex, 4] = data.CountPersRepeat;
                            objWorkSheet.Cells[currentIndex, 5] = data.CountMessages;
                            objWorkSheet.Cells[currentIndex, 6] = data.TotalCost;
                            objWorkSheet.Cells[currentIndex++, 7] = data.AccountingDocument;
                        }
                        currentIndex++;
                    }
                    objWorkSheet.Cells[currentIndex, 3] = themeData.TotalPersFirst;
                    objWorkSheet.Cells[currentIndex, 4] = themeData.TotalPersRepeat;
                }
                else
                {
                    foreach (var data in themeData.Data.OrderBy(x => x.Code))
                    {
                        objWorkSheet.Cells[currentIndex++, 7] = data.CountPersFirst;
                    }
                }
            }

            objWorkSheet.Cells[5, 4] = report.FilialName;
            objWorkSheet.Cells[6, 4] = header;
            objWorkSheet.Cells[181, 2] = CurrentUser.Director;
            objWorkSheet.Cells[181, 6] = DateTime.Today;
        }

        private void FillConsolidateReport262T1(External.CReport262Table1[] reportList)
        {
            int countReport = reportList.Length;
            int currentIndex = startPosition;
            CopyNullCells(objWorkSheet, countReport, startPosition);

            foreach (var data in reportList)
            {
                int startPpl = 3;
                int startInfo = 16;

                foreach (var count in data.ListOfCountPpl)
                {
                    objWorkSheet.Cells[currentIndex, startPpl++] = count;
                }
                foreach (var count in data.ListOfCountInform)
                {
                    objWorkSheet.Cells[currentIndex, startInfo++] = count;
                }

                objWorkSheet.Cells[currentIndex++, 1] = data.Filial;
            }
        }

        private void FillConsolidateReport262T2(External.CReport262Table2[] reportList,
            External.CReport262Table2[] yearReportList)
        {
            int countReport = reportList.Length;
            int currentIndex = startPosition;
            CopyNullCells(objWorkSheet, countReport, startPosition);

            foreach (var data in yearReportList)
            {
                var monthData = reportList.SingleOrDefault(x => x.Filial == data.Filial);
                objWorkSheet.Cells[currentIndex, 1] = data.Filial;
                objWorkSheet.Cells[currentIndex, 4] = monthData?.Data?.CountSms ?? 0;
                objWorkSheet.Cells[currentIndex, 5] = data.Data.CountSms;
                objWorkSheet.Cells[currentIndex, 6] = monthData?.Data?.CountPost ?? 0;
                objWorkSheet.Cells[currentIndex, 7] = data.Data.CountPost;
                objWorkSheet.Cells[currentIndex, 8] = monthData?.Data?.CountPhone ?? 0;
                objWorkSheet.Cells[currentIndex, 9] = data.Data.CountPhone;
                objWorkSheet.Cells[currentIndex, 10] = monthData?.Data?.CountMessengers ?? 0;
                objWorkSheet.Cells[currentIndex, 11] = data.Data.CountMessengers;
                objWorkSheet.Cells[currentIndex, 12] = monthData?.Data?.CountEmail ?? 0;
                objWorkSheet.Cells[currentIndex, 13] = data.Data.CountEmail;
                objWorkSheet.Cells[currentIndex, 14] = monthData?.Data?.CountAddress ?? 0;
                objWorkSheet.Cells[currentIndex, 15] = data.Data.CountAddress;
                objWorkSheet.Cells[currentIndex, 16] = monthData?.Data?.CountAnother ?? 0;
                objWorkSheet.Cells[currentIndex++, 17] = data.Data.CountAnother;
            }
        }

        private void FillConsolidateReport262T3(External.CReport262Table3[] reportList)
        {
            int countReport = reportList.Length;
            int currentIndex = startPosition;
            CopyNullCells(objWorkSheet, countReport, startPosition);

            foreach (var data in reportList)
            {
                objWorkSheet.Cells[currentIndex, 1] = data.Filial;
                objWorkSheet.Cells[currentIndex, 2] = data.Data.CountUnit;
                objWorkSheet.Cells[currentIndex, 3] = data.Data.CountUnitChild;
                objWorkSheet.Cells[currentIndex, 4] = data.Data.CountUnitWithSp;
                objWorkSheet.Cells[currentIndex, 5] = data.Data.CountUnitWithSpChild;
                objWorkSheet.Cells[currentIndex, 6] = data.Data.CountChannelSp;
                objWorkSheet.Cells[currentIndex, 7] = data.Data.CountChannelSpChild;
                objWorkSheet.Cells[currentIndex, 8] = data.Data.CountChannelPhone;
                objWorkSheet.Cells[currentIndex, 9] = data.Data.CountChannelPhoneChild;
                objWorkSheet.Cells[currentIndex, 10] = data.Data.CountChannelTerminal;
                objWorkSheet.Cells[currentIndex, 11] = data.Data.CountChannelTerminalChild;
                objWorkSheet.Cells[currentIndex, 12] = data.Data.CountChannelAnother;
                objWorkSheet.Cells[currentIndex++, 13] = data.Data.CountChannelAnotherChild;
            }
        }

        private string GetPhoneCode(string phone)
        {
            if (!string.IsNullOrEmpty(phone))
            {
                int index = phone.IndexOf(")");
                if (index > 0)
                {
                    return phone.Substring(1, CurrentUser.DirectorPhone.IndexOf(")") - 1);
                }

                return phone.Substring(0, 3);
            }
            return "";
        }

        private string GetPhoneNumber(string phone)
        {
            if (!string.IsNullOrEmpty(phone))
            {
                int index = phone.IndexOf(")");
                if (index > 0)
                {
                    return phone.Substring(CurrentUser.DirectorPhone.IndexOf(")") + 1);
                }

                return phone.Substring(3);
            }
            return "";
        }

        private void CopyNullCells(Worksheet objWorkSheet, int count, int position)
        {
            for (int k = 1; k <= count - 2; k++)
            {
                var r = objWorkSheet.Range[position + ":" + position, Type.Missing];
                r.Copy(Type.Missing);
                r = objWorkSheet.Range[Convert.ToString(k + position) + ":" + Convert.ToString(k + position), Type.Missing];
                r.Insert(XlInsertShiftDirection.xlShiftDown);
            }
        }

        private readonly List<ReportDictionary> iizlDictionaries = new List<ReportDictionary>
        {
            new ReportDictionary() { TableName =  "Согласие", StartRowIndex = 29 },
            new ReportDictionary() { TableName =  "Тема Д1", StartRowIndex = 43 },
            new ReportDictionary() { TableName =  "Тема Д2", StartRowIndex = 62 },
            new ReportDictionary() { TableName =  "Тема Д3", StartRowIndex = 81 },
            new ReportDictionary() { TableName =  "Тема Д4", StartRowIndex = 100 },
            new ReportDictionary() { TableName =  "Тема П", StartRowIndex = 119 },
            new ReportDictionary() { TableName =  "Тема С", StartRowIndex = 138 },
            new ReportDictionary() { TableName =  "Тема К", StartRowIndex = 150 },
            new ReportDictionary() { TableName =  "Тема О", StartRowIndex = 162 }
        };

        private readonly List<ReportDictionary> f294Dictionaries = new List<ReportDictionary>
        {
            new ReportDictionary()
            {
                TableName = "Таблица 1",
                StartRowIndex = 10,
                EndRowIndex = 32,
                RowNumIndex = 2,
                ColumnStartIndex = 4,
                Index = 2
            },
            new ReportDictionary()
            {
                TableName = "Таблица 2",
                StartRowIndex = 8,
                EndRowIndex = 27,
                RowNumIndex = 2,
                ColumnStartIndex = 4,
                Index = 2
            },
            new ReportDictionary()
            {
                TableName = "Таблица 3",
                StartRowIndex = 8,
                EndRowIndex = 31,
                RowNumIndex = 2,
                ColumnStartIndex = 6,
                Index = 16
            },
            new ReportDictionary()
            {
                TableName = "Таблица 4",
                StartRowIndex = 8,
                EndRowIndex = 23,
                RowNumIndex = 2,
                ColumnStartIndex = 6,
                Index = 16
            },
            new ReportDictionary()
            {
                TableName = "Таблица 5",
                StartRowIndex = 8,
                EndRowIndex = 18,
                RowNumIndex = 2,
                ColumnStartIndex = 6,
                Index = 16
            },
            new ReportDictionary()
            {
                TableName = "Таблица 6",
                StartRowIndex = 8,
                EndRowIndex = 18,
                RowNumIndex = 2,
                ColumnStartIndex = 6,
                Index = 13
            },
            new ReportDictionary()
            {
                TableName = "Таблица 7",
                StartRowIndex = 8,
                EndRowIndex = 31,
                RowNumIndex = 2,
                ColumnStartIndex = 4,
                Index = 2
            },
            new ReportDictionary()
            {
                TableName = "Таблица 8",
                StartRowIndex = 6,
                EndRowIndex = 13,
                RowNumIndex = 2,
                ColumnStartIndex = 4,
                Index = 1
            },
            new ReportDictionary()
            {
                TableName = "Таблица 9",
                StartRowIndex = 5,
                EndRowIndex = 15,
                RowNumIndex = 2,
                ColumnStartIndex = 4,
                Index = 2
            },
            new ReportDictionary()
            {
                TableName = "Эффективность",
                StartRowIndex = 5,
                EndRowIndex = 42,
                RowNumIndex = 3,
                ColumnStartIndex = 4,
                Index = 2
            }
        };

        private readonly List<ReportDictionary> pgDictionaries = new List<ReportDictionary>
        {
            new ReportDictionary() { TableName = "Таблица 1", StartRowIndex = 12, EndRowIndex = 70, Index = 1 },
            new ReportDictionary() { TableName = "Таблица 2", StartRowIndex = 7, EndRowIndex = 28, Index = 2 },
            new ReportDictionary() { TableName = "Таблица 3", StartRowIndex = 7, EndRowIndex = 35, Index = 3 },
            new ReportDictionary() { TableName = "Таблица 4", StartRowIndex = 7, EndRowIndex = 11, Index = 4 },
            new ReportDictionary() { TableName = "Таблица 5", StartRowIndex = 7, EndRowIndex = 34, Index = 5 },
            new ReportDictionary() { TableName = "Таблица 6", StartRowIndex = 7, EndRowIndex = 49, Index = 6 },
            new ReportDictionary() { TableName = "Таблица 8", StartRowIndex = 7, EndRowIndex = 72, Index = 8 },
            new ReportDictionary() { TableName = "Таблица 10", StartRowIndex = 7, EndRowIndex = 49, Index = 10 },
            new ReportDictionary() { TableName = "Таблица 11", StartRowIndex = 7, EndRowIndex = 27, Index = 11 },
            new ReportDictionary() { TableName = "Таблица 12", StartRowIndex = 7, EndRowIndex = 24, Index = 12 },
            new ReportDictionary() { TableName = "Таблица 13", StartRowIndex = 7, EndRowIndex = 21, Index = 13 },
        };

        private class ReportDictionary
        {
            public string TableName { get; set; }
            public int StartRowIndex { get; set; }
            public int EndRowIndex { get; set; }
            public int RowNumIndex { get; set; }
            public int ColumnStartIndex { get; set; }
            public int Index { get; set; }
        }
    }
}
