using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelFFOMSPersonnelCreator : ExcelBaseCreator<FFOMSPersonnel>
    {
        public ExcelFFOMSPersonnelCreator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.FFOMSPersonnel, header, filialName, true) { }

        protected override void FillReport(FFOMSPersonnel report, FFOMSPersonnel yearReport)
        {
            FillTable1(report);
        }

        private void FillTable1(FFOMSPersonnel report)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];

            int currentIndex = 4;
            foreach (var data in report.PersonnelT9)
            {
                ObjWorkSheet.Cells[currentIndex, 3] = data.FullTime;
                ObjWorkSheet.Cells[currentIndex++, 4] = data.Contract;
            }
        }
    }
}
