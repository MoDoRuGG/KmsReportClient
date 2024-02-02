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

    public class ExcelConsolidateQuantityAR : ExcelBaseCreator<ConsolidateQuantityAR[]>

    {

        public ExcelConsolidateQuantityAR(
     string filename,
     string header,
     string filialName) : base(filename, ExcelForm.ConsQuantityAR, header, filialName, false)
        {

        }


        protected override void FillReport(ConsolidateQuantityAR[] report, ConsolidateQuantityAR[] yearReport)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            CopyNullCells(ObjWorkSheet,report.Length/12,4);

            var filials = report.Select(x => x.RegionName).Distinct().OrderBy(x => x);
            int rowIndex = 4;
            foreach (var filial in filials)
            {
                var monthData = report.Where(x => x.RegionName == filial);
                ObjWorkSheet.Cells[rowIndex, 1] = filial;
                var columnIndex1 = 2;
                var columnIndex2 = 15;
                foreach (var md in monthData.OrderBy(x => x.Yymm))
                {

                    {
                        ObjWorkSheet.Cells[rowIndex, columnIndex1++] = md.Added;
                        ObjWorkSheet.Cells[rowIndex, columnIndex2++] = md.Removed;
                    }
                }
                rowIndex += 1;

            }
        }
    }
}
