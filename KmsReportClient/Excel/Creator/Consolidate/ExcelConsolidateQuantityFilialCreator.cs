using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelConsolidateQuantityFilialsCreator : ExcelBaseCreator<CReportQuantityFilial[]>
    {
        private const int StartPosition = 6;

        public ExcelConsolidateQuantityFilialsCreator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.ConsQuantityFilials, header, filialName, false) { }

        protected override void FillReport(CReportQuantityFilial[] report, CReportQuantityFilial[] yearReport)
        {
            int countReport = report.Length;
            int currentIndex = StartPosition;
            CopyNullCells(ObjWorkSheet, countReport+1, StartPosition);

            foreach (var data in report)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Data.Col_1;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Data.Col_2;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Data.Col_3;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Data.Col_4;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Data.Col_5;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Data.Col_6;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Data.Col_7;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Data.Col_8;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Data.Col_9;
                ObjWorkSheet.Cells[currentIndex, 11] = data.Data.Col_10;
                ObjWorkSheet.Cells[currentIndex, 12] = data.Data.Col_11;
                ObjWorkSheet.Cells[currentIndex, 13] = data.Data.Col_12;
                ObjWorkSheet.Cells[currentIndex, 14] = data.Data.Col_13;
                ObjWorkSheet.Cells[currentIndex, 15] = data.Data.Col_14;
                ObjWorkSheet.Cells[currentIndex, 16] = data.Data.Col_15;
                ObjWorkSheet.Cells[currentIndex, 17] = data.Data.Col_16;
                currentIndex++;
            }
        }
    }
}
