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
    class ExcelConsolidateCadri : ExcelBaseCreator<ConsolidateCadri[]>
    {
        public ExcelConsolidateCadri(
         string filename,
         string header,
         string filialName) : base(filename, ExcelForm.cadri, header, filialName, false) { }

        protected override void FillReport(ConsolidateCadri[] report, ConsolidateCadri[] yearReport)
        {
            FillCadri(report);
        }


        private void FillCadri(ConsolidateCadri[] reports)
        {
            var cardi = reports.Select(r => new { r.Filial, r.data }).ToList();

            int countReport = cardi.Count;
            int currentIndex = 6;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            CopyNullCellsCadri(ObjWorkSheet, countReport, currentIndex);

            int counter = 1;
            int currentIndexF = 5;
            foreach (var data in cardi)
            {
                //ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndexF, 1] = data.Filial;
                currentIndexF += 3;

                ObjWorkSheet.Cells[currentIndex, 3] = data.data.r1.IzNih1;
                ObjWorkSheet.Cells[currentIndex + 1, 3] = data.data.r1.IzNih2;
                
                ObjWorkSheet.Cells[currentIndex, 4] = data.data.r11.IzNih1;
                ObjWorkSheet.Cells[currentIndex + 1, 4] = data.data.r11.IzNih2;
        
                ObjWorkSheet.Cells[currentIndex, 5] = data.data.r111.IzNih1;
                ObjWorkSheet.Cells[currentIndex + 1, 5] = data.data.r111.IzNih2;


                ObjWorkSheet.Cells[currentIndex, 6] = data.data.r112.IzNih1;
                ObjWorkSheet.Cells[currentIndex + 1, 6] = data.data.r112.IzNih2;



                ObjWorkSheet.Cells[currentIndex, 7] = data.data.r113.IzNih1;
                ObjWorkSheet.Cells[currentIndex + 1, 7] = data.data.r113.IzNih2;



                ObjWorkSheet.Cells[currentIndex, 8] = data.data.r1131.IzNih1;
                ObjWorkSheet.Cells[currentIndex + 1, 8] = data.data.r1131.IzNih2;



                ObjWorkSheet.Cells[currentIndex, 9] = data.data.r11311.IzNih1;
                ObjWorkSheet.Cells[currentIndex + 1, 9] = data.data.r11311.IzNih2;


                ObjWorkSheet.Cells[currentIndex, 10] = data.data.r1132.IzNih1;
                ObjWorkSheet.Cells[currentIndex + 1, 10] = data.data.r1132.IzNih2;


                ObjWorkSheet.Cells[currentIndex, 11] = data.data.r11321.IzNih1;
                ObjWorkSheet.Cells[currentIndex + 1, 11] = data.data.r11321.IzNih2;



                ObjWorkSheet.Cells[currentIndex, 12] = data.data.r2.IzNih1;
                ObjWorkSheet.Cells[currentIndex + 1, 12] = data.data.r2.IzNih2;


                ObjWorkSheet.Cells[currentIndex, 13] = data.data.r21.IzNih1;
                ObjWorkSheet.Cells[currentIndex + 1, 13] = data.data.r21.IzNih2;


                ObjWorkSheet.Cells[currentIndex, 14] = data.data.r3.IzNih1;
                ObjWorkSheet.Cells[currentIndex + 1, 14] = data.data.r3.IzNih2;


                ObjWorkSheet.Cells[currentIndex, 15] = data.data.r31.IzNih1;
                ObjWorkSheet.Cells[currentIndex + 1, 15] = data.data.r31.IzNih2;


                ObjWorkSheet.Cells[currentIndex, 16] = data.data.r32.IzNih1;
                ObjWorkSheet.Cells[currentIndex + 1, 16] = data.data.r32.IzNih2;


                ObjWorkSheet.Cells[currentIndex, 17] = data.data.r33.IzNih1;
                ObjWorkSheet.Cells[currentIndex + 1, 17] = data.data.r33.IzNih2;


                ObjWorkSheet.Cells[currentIndex, 18] = data.data.r4.IzNih1;
                ObjWorkSheet.Cells[currentIndex + 1, 18] = data.data.r4.IzNih2;


                ObjWorkSheet.Cells[currentIndex, 19] = data.data.r41.IzNih1;
                ObjWorkSheet.Cells[currentIndex + 1, 19] = data.data.r41.IzNih2;


                ObjWorkSheet.Cells[currentIndex, 20] = data.data.r42.IzNih1;
                ObjWorkSheet.Cells[currentIndex + 1, 20] = data.data.r42.IzNih2;
                currentIndex+=3;





            }

        }

    }
}
