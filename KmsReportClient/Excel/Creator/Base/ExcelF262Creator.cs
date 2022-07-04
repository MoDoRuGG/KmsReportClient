using System;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Model.Enums;
using KmsReportClient.Support;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Base
{
    public class ExcelF262Creator : ExcelBaseCreator<Report262>
    {
        public ExcelF262Creator(
            string filename,
            ExcelForm reportName,
            string header,
            string filialName) : base(filename, reportName, header, filialName, false) { }

        protected override void FillReport(Report262 report, Report262 yearReport)
        {
            int i = 1;
            string reportMonths = "";
            string reportYear = "";
            if (report.Yymm.Length == 4)
            {
                reportMonths = YymmUtils.GetMonth(report.Yymm.Substring(2, 2));
                reportYear = report.Yymm.Substring(0, 2);
            }

            foreach (var themeData in report.ReportDataList.OrderBy(x => x.Theme))
            {
                switch (i)
                {
                    case 1:
                        FillTable1(themeData, reportMonths, reportYear, yearReport);
                        break;
                    case 2:
                        FillTable2(themeData, yearReport);
                        break;
                    default:
                        FillTable3(themeData, reportMonths, reportYear);
                        break;
                }

                if (i < 3)
                {
                    ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[++i];
                }
            }
        }

        private void FillTable1(Report262Dto themeData, string reportMonths, string reportYear, Report262 yearReport)
        {
            int currentIndex = 20;
            foreach (var data in themeData.Data.OrderBy(x => x.RowNum))
            {
                var yearTheme = yearReport.ReportDataList
                    .Where(x => x.Theme == themeData.Theme)
                    .SelectMany(x => x.Data);
                var yearData = yearTheme.Single(x => x.RowNum == data.RowNum);
                ObjWorkSheet.Cells[currentIndex, 80] = data.CountPpl;
                ObjWorkSheet.Cells[currentIndex++, 93] = yearData.CountPpl;
            }

            ObjWorkSheet.Cells[12, 19] = FilialName;
            ObjWorkSheet.Cells[10, 40] = reportMonths;
            ObjWorkSheet.Cells[10, 62] = reportYear;
        }

        private void FillTable2(Report262Dto themeData, Report262 yearReport)
        {
            if (themeData.Data.Length == 0)
            {
                return;
            }

            var data = themeData.Data[0];
            var yearTheme = yearReport.ReportDataList
                .Where(x => x.Theme == themeData.Theme)
                .SelectMany(x => x.Data);
            var yearData = yearTheme.ToArray()[0];
            ObjWorkSheet.Cells[7, 56] = data.CountSms;
            ObjWorkSheet.Cells[7, 65] = yearData.CountSms;
            ObjWorkSheet.Cells[7, 72] = data.CountPost;
            ObjWorkSheet.Cells[7, 81] = yearData.CountPost;
            ObjWorkSheet.Cells[7, 88] = data.CountPhone;
            ObjWorkSheet.Cells[7, 97] = yearData.CountPhone;
            ObjWorkSheet.Cells[7, 104] = data.CountMessengers;
            ObjWorkSheet.Cells[7, 113] = yearData.CountMessengers;
            ObjWorkSheet.Cells[7, 120] = data.CountEmail;
            ObjWorkSheet.Cells[7, 129] = yearData.CountEmail;
            ObjWorkSheet.Cells[7, 136] = data.CountAddress;
            ObjWorkSheet.Cells[7, 145] = yearData.CountAddress;
            ObjWorkSheet.Cells[7, 152] = data.CountAnother;
            ObjWorkSheet.Cells[7, 161] = yearData.CountAnother;

            ObjWorkSheet.Cells[11, 52] = CurrentUser.Director;
            ObjWorkSheet.Cells[14, 58] = GetPhoneCode(CurrentUser.DirectorPhone);
            ObjWorkSheet.Cells[14, 67] = GetPhoneNumber(CurrentUser.DirectorPhone);

            ObjWorkSheet.Cells[18, 52] = CurrentUser.UserName;
            ObjWorkSheet.Cells[21, 58] = GetPhoneCode(CurrentUser.Phone);
            ObjWorkSheet.Cells[21, 67] = GetPhoneNumber(CurrentUser.Phone);

            ObjWorkSheet.Cells[21, 6] = CurrentUser.Email;

            var date = DateTime.Today;
            ObjWorkSheet.Cells[14, 8] = date.Day;
            ObjWorkSheet.Cells[14, 15] = date.ToString("MMMM");
            ObjWorkSheet.Cells[14, 37] = date.ToString("yy");
        }

        private void FillTable3(Report262Dto themeData, string reportMonths, string reportYear)
        {
            var date = DateTime.Today;
            ObjWorkSheet.Cells[8, 72] = reportMonths;
            ObjWorkSheet.Cells[8, 94] = reportYear;

            ObjWorkSheet.Cells[21, 52] = CurrentUser.Director;
            ObjWorkSheet.Cells[24, 58] = GetPhoneCode(CurrentUser.DirectorPhone);
            ObjWorkSheet.Cells[24, 67] = GetPhoneNumber(CurrentUser.DirectorPhone);

            ObjWorkSheet.Cells[28, 52] = CurrentUser.UserName;
            ObjWorkSheet.Cells[31, 58] = GetPhoneCode(CurrentUser.Phone);
            ObjWorkSheet.Cells[31, 67] = GetPhoneNumber(CurrentUser.Phone);
            ObjWorkSheet.Cells[31, 6] = CurrentUser.Email;

            ObjWorkSheet.Cells[24, 8] = date.Day;
            ObjWorkSheet.Cells[24, 15] = date.ToString("MMMM");
            ObjWorkSheet.Cells[24, 37] = date.ToString("yy");

            int position = 16;
            CopyNullCells(ObjWorkSheet, themeData.Table3.Length, position);
            foreach (var data in themeData.Table3)
            {
                ObjWorkSheet.Cells[position, 1] = data.Mo;
                ObjWorkSheet.Cells[position, 55] = data.CountUnit;
                ObjWorkSheet.Cells[position, 66] = data.CountUnitChild;
                ObjWorkSheet.Cells[position, 75] = data.CountUnitWithSp;
                ObjWorkSheet.Cells[position, 90] = data.CountUnitWithSpChild;
                ObjWorkSheet.Cells[position, 103] = data.CountChannelSp;
                ObjWorkSheet.Cells[position, 111] = data.CountChannelSpChild;
                ObjWorkSheet.Cells[position, 119] = data.CountChannelPhone;
                ObjWorkSheet.Cells[position, 127] = data.CountChannelPhoneChild;
                ObjWorkSheet.Cells[position, 136] = data.CountChannelTerminal;
                ObjWorkSheet.Cells[position, 144] = data.CountChannelTerminalChild;
                ObjWorkSheet.Cells[position, 152] = data.CountChannelAnother;
                ObjWorkSheet.Cells[position, 160] = data.CountChannelAnotherChild;
                position++;
            }
        }
    }
}
