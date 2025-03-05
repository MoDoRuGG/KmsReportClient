using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelFFOMSTargetedExpCreator : ExcelBaseCreator<FFOMSTargetedExp>
    {
        public ExcelFFOMSTargetedExpCreator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.FFOMSTargetedExp, header, filialName, true) { }

        protected override void FillReport(FFOMSTargetedExp report, FFOMSTargetedExp yearReport)
        {
            FillTable1(report);
            FillTable2(report);
            FillTable3(report);
        }

        private void FillTable1(FFOMSTargetedExp report)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];

            int currentIndex = 3;
            foreach (var data in report.MEE)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = data.Target;
            }
        }


        private void FillTable2(FFOMSTargetedExp report)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[2];
            int currentIndex = 3;
            foreach (var data in report.EKMP)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = data.Target;
            }
        }

        private void FillTable3(FFOMSTargetedExp report)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[3];
            int currentIndex = 3;
            foreach (var data in report.MD_EKMP)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = data.Target;

            }
        }
    }
}
