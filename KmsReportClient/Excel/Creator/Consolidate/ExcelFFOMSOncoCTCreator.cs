using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelFFOMSOncoCTCreator : ExcelBaseCreator<FFOMSOncoCT>
    {
        public ExcelFFOMSOncoCTCreator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.FFOMSOncoCT, header, filialName, true) { }

        protected override void FillReport(FFOMSOncoCT report, FFOMSOncoCT yearReport)
        {
            FillTable1(report);
        }

        private void FillTable1(FFOMSOncoCT report)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];

            foreach (var data in report.OncoCT_MEE)
            {
                ObjWorkSheet.Cells[2, 2] = data.Target;
                ObjWorkSheet.Cells[2, 3] = data.Target;
            }
        }
    }
}
