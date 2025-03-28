using System.Collections.Generic;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.Excel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelFFOMSViolEKMPCreator : ExcelBaseCreator<FFOMSViolEKMP>
    {
        public ExcelFFOMSViolEKMPCreator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.ViolEKMP, header, filialName, true) { }



        protected override void FillReport(FFOMSViolEKMP report, FFOMSViolEKMP yearReport)
        {
            int currentIndex = 3;
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            foreach (var data in report.DataViolEKMP)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = data.Count;
            }
        }
    }
}
