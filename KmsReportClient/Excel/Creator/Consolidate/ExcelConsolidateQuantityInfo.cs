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

    public class ExcelConsolidateQuantityInfo : ExcelBaseCreator<ConsolidateQuantityInfo[]>

    {

        public ExcelConsolidateQuantityInfo(
     string filename,
     string header,
     string filialName) : base(filename, ExcelForm.ConsQuantityInformation, header, filialName, false)
        {

        }


        protected override void FillReport(ConsolidateQuantityInfo[] report, ConsolidateQuantityInfo[] yearReport)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            CopyNullCells(ObjWorkSheet,report.Length/12+1,4);

            var filials = report.Select(x => x.RegionName).Distinct().OrderBy(x => x);
            int rowIndex = 4;
            foreach (var filial in filials)
            {
                var monthData = report.Where(x => x.RegionName == filial);
                ObjWorkSheet.Cells[rowIndex, 1] = filial;
                var columnIndex1 = 2;
                var columnIndex2 = 14;
                var columnIndex3 = 26;
                var columnIndex4 = 50;
                var columnIndex5 = 62;
                var columnIndex6 = 74;
                var columnIndex7 = 86;
                var columnIndex8 = 98;
                var columnIndex9 = 110;

                foreach (var md in monthData.OrderBy(x => x.Yymm))
                {
                    {
                        ObjWorkSheet.Cells[rowIndex, columnIndex1++] = md.Fact;
                        ObjWorkSheet.Cells[rowIndex, columnIndex2++] = md.Plan;
                        ObjWorkSheet.Cells[rowIndex, columnIndex3++] = md.Added-md.Plan;
                        ObjWorkSheet.Cells[rowIndex, columnIndex4++] = md.Col_3 + md.Col_7 + md.Col_15;
                        ObjWorkSheet.Cells[rowIndex, columnIndex5++] = md.Col_8;
                        ObjWorkSheet.Cells[rowIndex, columnIndex6++] = md.Col_10+md.Col_12;
                        ObjWorkSheet.Cells[rowIndex, columnIndex7++] = md.Col_10;
                        ObjWorkSheet.Cells[rowIndex, columnIndex8++] = md.Col_12;
                        ObjWorkSheet.Cells[rowIndex, columnIndex9++] = md.Fact-md.Col_1;
                    }
                }
                rowIndex += 1;

            }
        }
    }
}
