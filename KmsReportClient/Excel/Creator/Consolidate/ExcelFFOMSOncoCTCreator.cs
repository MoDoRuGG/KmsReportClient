using System.Linq;
using System.Windows.Forms;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelFFOMSOncoCTCreator : ExcelBaseCreator<FFOMSOncoCT[]>
    {

        private string _yymm;
        public ExcelFFOMSOncoCTCreator(
            string filename,
            string header,
            string filialName, string yymm) : base(filename, ExcelForm.FFOMSOncoCT, header, filialName, false) { _yymm = yymm; }

        protected override void FillReport(FFOMSOncoCT[] reports, FFOMSOncoCT[] yearReports)
        {
            var expertises = reports.Select(r => new { r.Filial, r.OncoCT_MEE }).ToList();
            int countReport = expertises.Count;
            int currentIndex = 2;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            CopyNullCells(ObjWorkSheet, countReport, 2);


            foreach (var data in expertises)

            {
                ObjWorkSheet.Cells[currentIndex, 1] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 2] = data.OncoCT_MEE.Target;
                ObjWorkSheet.Cells[currentIndex++, 3] = data.OncoCT_MEE.Target;
            }
        }
    }
}
