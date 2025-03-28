using System.Collections.Generic;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.Excel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelFFOMSVerifyPlanCreator : ExcelBaseCreator<FFOMSVerifyPlan>
    {
        public ExcelFFOMSVerifyPlanCreator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.VerifyPlan, header, filialName, true) { }



        protected override void FillReport(FFOMSVerifyPlan report, FFOMSVerifyPlan yearReport)
        {
            int currentIndex = 2;
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            foreach (var data in report.DataVerifyPlan)
            {
                ObjWorkSheet.Cells[currentIndex++, 4] = data.Count;
            }
        }
    }
}
