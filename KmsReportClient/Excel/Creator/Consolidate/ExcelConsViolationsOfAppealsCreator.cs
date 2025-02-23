using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelConsViolationsOfAppealsCreator : ExcelBaseCreator<ViolationsOfAppeals>
    {
        public ExcelConsViolationsOfAppealsCreator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.ViolationsOfAppeals, header, filialName, true) { }

        protected override void FillReport(ViolationsOfAppeals report, ViolationsOfAppeals yearReport)
        {
            FillTable1(report);
            FillTable2(report);
            FillTable3(report);
        }

        private void FillTable1(ViolationsOfAppeals report)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            ObjWorkSheet.Range["B1","C1"].Value = $"Количество обоснованных жалоб застрахованных лиц  в {FilialName} за 2024 год";

            int currentIndex = 2;
            foreach (var treatment in report.T1)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = treatment.Oral+treatment.Written;

            }
        }


        private void FillTable2(ViolationsOfAppeals report)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[2];
            ObjWorkSheet.Range["B1", "C1"].Value = $"Количество страховых случаев, подвергшихся ЭКМП, проведенным по жалобам от застрахованных лиц или их представителей в {FilialName} за 2024 год";
            int currentIndex = 2;
            foreach (var expertise in report.T2)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = expertise.Target + expertise.Plan;
            }
        }

        private void FillTable3(ViolationsOfAppeals report)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[3];
            ObjWorkSheet.Range["B1", "C1"].Value = $"Количество страховых случаев, подвергшихся МЭЭ, проведенным по жалобам от застрахованных лиц или их представителей в {FilialName} за 2024 год";
            int currentIndex = 2;
            foreach (var expertise in report.T3)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = expertise.Target + expertise.Plan;

            }
        }
    }
}
