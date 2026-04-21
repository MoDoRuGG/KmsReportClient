using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    public class ExcelConsolidateOpedUnplQCreator : ExcelBaseCreator<ConsolidateOpedUnplQ[]>
    {

        public ExcelConsolidateOpedUnplQCreator(
         string filename,
         string header,
         string filialName) : base(filename, ExcelForm.consOpedUnplQ, header, filialName, false)
        {

        }

        protected override void FillReport(ConsolidateOpedUnplQ[] report, ConsolidateOpedUnplQ[] yearReport)
        {
            int row = 4;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            CopyNullCells(ObjWorkSheet, report.Length, row);

            int counter = 1;
            foreach (var regionData in report)
            {
                ObjWorkSheet.Cells[row, 1] = counter++;
                ObjWorkSheet.Cells[row, 2] = regionData.Region;

                ObjWorkSheet.Cells[row, 3] = regionData.LethalPlan;
                ObjWorkSheet.Cells[row, 4] = regionData.LethalFact;

                ObjWorkSheet.Cells[row, 6] = regionData.PovtorPlan;
                ObjWorkSheet.Cells[row, 7] = regionData.PovtorFact;

                ObjWorkSheet.Cells[row, 9] = regionData.OncoPlan;
                ObjWorkSheet.Cells[row, 10] = regionData.OncoFact;

                ObjWorkSheet.Cells[row, 12] = regionData.EcoPlan;
                ObjWorkSheet.Cells[row, 13] = regionData.EcoFact;

                ObjWorkSheet.Cells[row, 15] = regionData.Notes;

                row++;
            }

        }
    }
}
