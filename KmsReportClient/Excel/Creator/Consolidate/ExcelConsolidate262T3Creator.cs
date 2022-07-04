using KmsReportClient.External;
using KmsReportClient.Model.Enums;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelConsolidate262T3Creator : ExcelBaseCreator<CReport262Table3[]>
    {
        private const int StartPosition = 5;

        public ExcelConsolidate262T3Creator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.C262T3, header, filialName, false) { }

        protected override void FillReport(CReport262Table3[] reports, CReport262Table3[] yearReport)
        {
            int countReport = reports.Length;
            int currentIndex = StartPosition;
            CopyNullCells(ObjWorkSheet, countReport, StartPosition);

            foreach (var data in reports)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Data.CountUnit;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Data.CountUnitChild;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Data.CountUnitWithSp;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Data.CountUnitWithSpChild;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Data.CountChannelSp;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Data.CountChannelSpChild;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Data.CountChannelPhone;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Data.CountChannelPhoneChild;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Data.CountChannelTerminal;
                ObjWorkSheet.Cells[currentIndex, 11] = data.Data.CountChannelTerminalChild;
                ObjWorkSheet.Cells[currentIndex, 12] = data.Data.CountChannelAnother;
                ObjWorkSheet.Cells[currentIndex++, 13] = data.Data.CountChannelAnotherChild;
            }
        }
    }
}
