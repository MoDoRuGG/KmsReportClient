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
    public class ExcelConsolidateOpedQCreator : ExcelBaseCreator<ConsolidateOpedQ[]>
    {
        
        public ExcelConsolidateOpedQCreator(
         string filename,
         string header,
         string filialName) : base(filename, ExcelForm.consOpedQ, header, filialName, false)
        {
          
        }

        protected override void FillReport(ConsolidateOpedQ[] report, ConsolidateOpedQ[] yearReport)
        {
            int row = 4;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            CopyNullCells(ObjWorkSheet, report.Length, row);

            int counter = 1;
            foreach(var regionData in report)
            {
                ObjWorkSheet.Cells[row, 1] = counter++;
                ObjWorkSheet.Cells[row, 2] = regionData.Region;

                ObjWorkSheet.Cells[row, 3] = regionData.MeePovtorPlan;
                ObjWorkSheet.Cells[row, 4] = regionData.MeePovtorFact;

                ObjWorkSheet.Cells[row, 6] = regionData.MeeOnkoPlan;
                ObjWorkSheet.Cells[row, 7] = regionData.MeeOnkoFact;

                ObjWorkSheet.Cells[row, 9] = regionData.EkmpLetalPlan;
                ObjWorkSheet.Cells[row, 10] = regionData.EkmpLetalFact;

                ObjWorkSheet.Cells[row, 12] = regionData.Notes;
                ObjWorkSheet.Cells[row, 13] = regionData.NotesGoodReason;

                row++;
            }
        
        }
    }
}
