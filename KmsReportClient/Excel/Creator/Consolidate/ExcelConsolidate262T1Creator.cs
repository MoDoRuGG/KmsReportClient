using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelConsolidate262T1Creator : ExcelBaseCreator<CReport262Table1[]>
    {
        private const int StartPosition = 5;

        public ExcelConsolidate262T1Creator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.C262T1, header, filialName, false) { }

        protected override void FillReport(CReport262Table1[] report, CReport262Table1[] yearReport)
        {
            int countReport = report.Length;
            int currentIndex = StartPosition;
            CopyNullCells(ObjWorkSheet, countReport, StartPosition);

            foreach (var data in report)
            {
                int startPpl = 3;
                int startInfo = 16;

                foreach (var count in data.ListOfCountPpl.OrderBy(x => x.yymm))
                {
                    ObjWorkSheet.Cells[currentIndex, startPpl++] = count.Value;
                }
                foreach (var count in data.ListOfCountInform.OrderBy(x => x.yymm))
                {
                    ObjWorkSheet.Cells[currentIndex, startInfo++] = count.Value;
                }

                ObjWorkSheet.Cells[currentIndex++, 1] = data.Filial;
            }
        }
    }
}
