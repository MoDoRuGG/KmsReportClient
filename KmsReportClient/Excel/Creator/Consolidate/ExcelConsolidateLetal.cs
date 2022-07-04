using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    public class ExcelConsolidateLetal : ExcelBaseCreator<ConsolidateLetal[]>
    {
        private const int StartPosition = 14;

        public ExcelConsolidateLetal(
           string filename,
           string header,
           string filialName) : base(filename, ExcelForm.letal, header, filialName, false) { }


        protected override void FillReport(ConsolidateLetal[] report, ConsolidateLetal[] yearReport)
        {
            FillLetal(report);

        }

        public void FillLetal(ConsolidateLetal[] reports)
        {
            var letal = reports.Select(r => new { r.Filial, r.Data }).ToList();

            int countReport = letal.Count;
            int currentIndex = StartPosition;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            CopyNullCells(ObjWorkSheet, countReport, StartPosition);

            int counter = 1;
            foreach (var data in letal)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Data.r1;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Data.r1_1;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Data.r1_2;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Data.r121;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Data.r2;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Data.r3;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Data.r31;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Data.r311;
                ObjWorkSheet.Cells[currentIndex, 11] = data.Data.r3111;
                ObjWorkSheet.Cells[currentIndex, 12] = data.Data.r3112;
                ObjWorkSheet.Cells[currentIndex, 13] = data.Data.r3113;
                ObjWorkSheet.Cells[currentIndex, 14] = data.Data.r3114;
                ObjWorkSheet.Cells[currentIndex, 15] = data.Data.r32;
                ObjWorkSheet.Cells[currentIndex, 16] = data.Data.r33;
                ObjWorkSheet.Cells[currentIndex, 17] = data.Data.r4;
                ObjWorkSheet.Cells[currentIndex, 18] = data.Data.r5;
                ObjWorkSheet.Cells[currentIndex, 19] = data.Data.r6;
                ObjWorkSheet.Cells[currentIndex, 20] = data.Data.r7;
                ObjWorkSheet.Cells[currentIndex, 21] = data.Data.r8;
                ObjWorkSheet.Cells[currentIndex, 22] = data.Data.r9;
                ObjWorkSheet.Cells[currentIndex, 23] = data.Data.r10;
                ObjWorkSheet.Cells[currentIndex, 24] = data.Data.r11;
                ObjWorkSheet.Cells[currentIndex, 25] = data.Data.r12;
                ObjWorkSheet.Cells[currentIndex++, 26] = data.Data.r13;


            }
        }

    }


}
