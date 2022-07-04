using KmsReportClient.Excel.Creator.Base;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelConsolidateFilial294Creator : ExcelBaseCreator<Report294[]>
    {
        public ExcelConsolidateFilial294Creator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.CFilial294, header, filialName, false) { }

        protected override void FillReport(Report294[] reportList, Report294[] yearReport)
        {
            var excel294 = new ExcelF294Creator(Filename, ReportName, Header, FilialName);

            int month = 0;
            foreach (var report in reportList)
            {
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
                excel294.FillTables(report, null, month, ObjWorkSheet, ObjWorkBook);
                month++;
            }
        }
    }
}
