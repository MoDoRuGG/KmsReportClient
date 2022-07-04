using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelConsolidateCpnpMonthCreator : ExcelBaseCreator<ConsolidateCpnpM[]>
    {
        private const int StartPosition = 9;
        string headerReport;

        public ExcelConsolidateCpnpMonthCreator(
           string filename,
           string header,
           string filialName, string mm, int year) : base(filename, ExcelForm.CСnpnM, header, filialName, false) { headerReport = $"За {mm} {year} года"; }

        protected override void FillReport(ConsolidateCpnpM[] report, ConsolidateCpnpM[] yearReport)
        {
            int countReport = report.Length;
            int currentIndex = StartPosition;

            int number = 1;

            ObjWorkSheet.Cells[4, 9] = headerReport;

            CopyNullCells(ObjWorkSheet, countReport, StartPosition);

            foreach (var data in report)
            {


                ObjWorkSheet.Cells[currentIndex, 1] = number++;

                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
        
                ObjWorkSheet.Cells[currentIndex, 3] = data.CountAll;
             
                ObjWorkSheet.Cells[currentIndex++, 4] = data.CountReason;
            }

        }
    }
}
