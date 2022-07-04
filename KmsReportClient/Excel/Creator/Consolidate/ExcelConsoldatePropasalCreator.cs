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
    public class ExcelConsoldatePropasalCreator : ExcelBaseCreator<ConsolidateProposal[]>
    {
        public ExcelConsoldatePropasalCreator(
         string filename,
         string header,
         string filialName) : base(filename, ExcelForm.consProposal, header, filialName, false) { }



        protected override void FillReport(ConsolidateProposal[] report, ConsolidateProposal[] yearReport)
        {
            int rowIndex = 4;
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            CopyNullCells(ObjWorkSheet, report.Length, rowIndex);

            foreach (var rep in report)
            {
                ObjWorkSheet.Cells[rowIndex, 1] = rep.RegionName;
                ObjWorkSheet.Cells[rowIndex, 2] = rep.CountMoCheck;
                ObjWorkSheet.Cells[rowIndex, 3] = rep.CountMoCheckWithDefect;
                ObjWorkSheet.Cells[rowIndex, 4] = rep.CountProporsals;
                ObjWorkSheet.Cells[rowIndex, 5] = rep.CountProporsalsWithDefect;
                if (!string.IsNullOrEmpty(rep.Notes))
                    ObjWorkSheet.Cells[rowIndex, 7] = rep.Notes;

                rowIndex++;
            }
        }
    }
}
