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

    public class ExcelConsolidateQuantityQ : ExcelBaseCreator<ConsolidateQuantityQ[]>

    {

        public ExcelConsolidateQuantityQ(
     string filename,
     string header,
     string filialName) : base(filename, ExcelForm.ConsQuantityQ, header, filialName, false)
        {

        }


        protected override void FillReport(ConsolidateQuantityQ[] report, ConsolidateQuantityQ[] yearReport)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            CopyNullCells(ObjWorkSheet,report.Length-1,4);

            var stYymm = report.First().Yymm.Substring(0,2);
            var finYymm = report.First().Yymm.Substring(2, 2);
            var filials = report.Select(x => x.RegionName).Distinct().OrderBy(x => x);
            int rowIndex = 3;
            ObjWorkSheet.Cells[2, 2] = "Численность на 01.01.20" + stYymm;
            ObjWorkSheet.Cells[2, 3] = "Численность на 01." +  finYymm + ".20" + stYymm;
            foreach (var filial in filials)
            {
                var Data = report.Where(x => x.RegionName == filial);
                ObjWorkSheet.Cells[rowIndex, 1] = filial;
                var columnIndex = 2;
                foreach (var md in Data.OrderBy(x => x.Yymm))
                {

                    {
                        ObjWorkSheet.Cells[rowIndex, columnIndex++] = md.c2;
                        ObjWorkSheet.Cells[rowIndex, columnIndex++] = md.c3;
                        ObjWorkSheet.Cells[rowIndex, columnIndex++] = md.c4;
                        ObjWorkSheet.Cells[rowIndex, columnIndex++] = md.c5;
                        ObjWorkSheet.Cells[rowIndex, columnIndex++] = md.c6;
                        ObjWorkSheet.Cells[rowIndex, columnIndex++] = md.c7;
                        ObjWorkSheet.Cells[rowIndex, columnIndex++] = md.c8;
                        ObjWorkSheet.Cells[rowIndex, columnIndex++] = md.c9;
                    }
                }
                rowIndex += 1;

            }
        }
    }
}
