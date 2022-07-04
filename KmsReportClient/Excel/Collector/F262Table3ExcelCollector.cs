using System;
using System.Collections.Generic;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Support;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Collector
{
    class F262Table3ExcelCollector : ExcelBaseCollector
    {
        protected override void FillReport(string form, AbstractReport destReport, AbstractReport srcReport)
        {
            var srcReportData = (destReport as Report262)?.ReportDataList.Single(r => r.Theme == form) ?? 
                                throw new Exception($"Can't find srcReportDataList for form = {form}");
            var destReportData = (srcReport as Report262)?.ReportDataList.Single(r => r.Theme == form) ?? 
                                 throw new Exception($"Can't find destReportDataList for form = {form}");
            srcReportData.Table3 = destReportData.Table3;
        }

        protected override AbstractReport CollectReportData(string form)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            int lastRow = GetLastRow();
            var list = new List<Report262Table3Data>();
            for (int i = 17; i <= lastRow; i++)
            {
                var data = CollectElement(i);
                list.Add(data);
            }

            var report = new Report262 { ReportDataList = new Report262Dto[1] };
            report.ReportDataList[0] = new Report262Dto
            {
                Table3 = list.ToArray(),
                Theme = form
            };

            return report;
        }

        private Report262Table3Data CollectElement(int i) => new Report262Table3Data
        {
            Mo = ObjWorkSheet.Cells[i, 2].Text,
            CountUnit = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 3].Text),
            CountUnitChild = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 4].Text),
            CountUnitWithSp = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 5].Text),
            CountUnitWithSpChild = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 6].Text),
            CountChannelSp = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 7].Text),
            CountChannelSpChild = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 8].Text),
            CountChannelPhone = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 9].Text),
            CountChannelPhoneChild = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 10].Text),
            CountChannelTerminal = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 11].Text),
            CountChannelTerminalChild = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 12].Text),
            CountChannelAnother = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 13].Text),
            CountChannelAnotherChild = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 14].Text)
        };
    }
}