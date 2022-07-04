using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelConsolidateCnpnCreator : ExcelBaseCreator<ConsolidateCpnp[]>
    {
        private const int StartPosition = 7;
        string headerReport;

        
        public ExcelConsolidateCnpnCreator(
            string filename,
            string header,
            string filialName,int q,int year) : base(filename, ExcelForm.CСnpnQ, header, filialName, false) { headerReport = $"За {q} квартал {year} года"; }

        protected override void FillReport(ConsolidateCpnp[] report, ConsolidateCpnp[] yearReport)
        {
            int countReport = report.Length;
            int currentIndex = StartPosition;

            int number = 1;

            ObjWorkSheet.Cells[3, 12] = headerReport;

            CopyNullCells(ObjWorkSheet, countReport, StartPosition);

            foreach (var data in report)
            {
               

                ObjWorkSheet.Cells[currentIndex, 1] = number++;

                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;

                ObjWorkSheet.Cells[currentIndex, 3] = data.CountPretrial;

                ObjWorkSheet.Cells[currentIndex, 4] = data.CountAll;

                ObjWorkSheet.Cells[currentIndex, 6] = data.NormativRegionCpnp;

                ObjWorkSheet.Cells[currentIndex++, 8] = data.NormativFederalCpnp;
            }

        }
    }
}
