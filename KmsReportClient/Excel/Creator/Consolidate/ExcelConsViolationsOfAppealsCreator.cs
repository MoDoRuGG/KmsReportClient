using System;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelConsViolationsOfAppealsCreator : ExcelBaseCreator<ViolationsOfAppeals>
    {
        public string period = "";
        public ExcelConsViolationsOfAppealsCreator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.ViolationsOfAppeals, header, filialName, true) { }

        protected override void FillReport(ViolationsOfAppeals report, ViolationsOfAppeals yearReport)
        {


            
            string year = (2000 + Convert.ToInt32(report.Yymm.Substring(0, 2))).ToString();
            string month = report.Yymm.Substring(2, 2);
            switch (month)
            {
                case "03":
                    period = $"1 квартал {year} года";
                    break;
                case "06":
                    period = $"1 полугодие {year} года";
                    break;
                case "09":
                    period = $"9 месяцев {year} года";
                    break;
                case "12":
                    period = $"{year} год";
                    break;


            }

            FillTable1(report);
            FillTable2(report);
            FillTable3(report);
        }

        private void FillTable1(ViolationsOfAppeals report)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            ObjWorkSheet.Range["B1","C1"].Value = $"Количество обоснованных жалоб застрахованных лиц  в {FilialName} за {period}";

            int currentIndex = 2;
            foreach (var treatment in report.T1)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = treatment.Oral+treatment.Written;

            }
        }


        private void FillTable2(ViolationsOfAppeals report)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[2];
            ObjWorkSheet.Range["B1", "C1"].Value = $"Количество страховых случаев, подвергшихся ЭКМП, проведенным по жалобам от застрахованных лиц или их представителей в {FilialName} за {period}";
            int currentIndex = 2;
            foreach (var expertise in report.T2)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = expertise.Target + expertise.Plan;
            }
        }

        private void FillTable3(ViolationsOfAppeals report)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[3];
            ObjWorkSheet.Range["B1", "C1"].Value = $"Количество страховых случаев, подвергшихся МЭЭ, проведенным по жалобам от застрахованных лиц или их представителей в {FilialName} за {period}";
            int currentIndex = 2;
            foreach (var expertise in report.T3)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = expertise.Target + expertise.Plan;

            }
        }
    }
}
