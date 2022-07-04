using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelConsolidate262T2Creator : ExcelBaseCreator<CReport262Table2[]>
    {
        private const int StartPosition = 5;

        public ExcelConsolidate262T2Creator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.C262T2, header, filialName, false) { }

        protected override void FillReport(CReport262Table2[] reports, CReport262Table2[] yearReport)
        {
            int countReport = reports.Length;
            int currentIndex = StartPosition;
            CopyNullCells(ObjWorkSheet, countReport, StartPosition);

            foreach (var data in yearReport)
            {
                var monthData = reports.SingleOrDefault(x => x.Filial == data.Filial);
                ObjWorkSheet.Cells[currentIndex, 1] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 4] = monthData?.Data?.CountSms ?? 0;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Data.CountSms;
                ObjWorkSheet.Cells[currentIndex, 6] = monthData?.Data?.CountPost ?? 0;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Data.CountPost;
                ObjWorkSheet.Cells[currentIndex, 8] = monthData?.Data?.CountPhone ?? 0;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Data.CountPhone;
                ObjWorkSheet.Cells[currentIndex, 10] = monthData?.Data?.CountMessengers ?? 0;
                ObjWorkSheet.Cells[currentIndex, 11] = data.Data.CountMessengers;
                ObjWorkSheet.Cells[currentIndex, 12] = monthData?.Data?.CountEmail ?? 0;
                ObjWorkSheet.Cells[currentIndex, 13] = data.Data.CountEmail;
                ObjWorkSheet.Cells[currentIndex, 14] = monthData?.Data?.CountAddress ?? 0;
                ObjWorkSheet.Cells[currentIndex, 15] = data.Data.CountAddress;
                ObjWorkSheet.Cells[currentIndex, 16] = monthData?.Data?.CountAnother ?? 0;
                ObjWorkSheet.Cells[currentIndex++, 17] = data.Data.CountAnother;
            }
        }
    }
}
