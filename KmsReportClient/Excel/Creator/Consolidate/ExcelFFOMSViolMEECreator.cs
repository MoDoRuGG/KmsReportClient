using System.Collections.Generic;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.Excel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelFFOMSViolMEECreator : ExcelBaseCreator<FFOMSViolMEE>
    {
        public ExcelFFOMSViolMEECreator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.ViolMEE, header, filialName, true) { }



        protected override void FillReport(FFOMSViolMEE report, FFOMSViolMEE yearReport)
        {
            int currentIndex = 3;
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            foreach (var data in report.DataViolMEE)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = data.Count;
            }
        }
    }
}
