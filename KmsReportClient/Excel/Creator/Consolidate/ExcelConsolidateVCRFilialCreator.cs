using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelConsolidateVCRFilialCreator : ExcelBaseCreator<CReportVCRFilial[]>
    {
        private const int StartPosition = 7;

        public ExcelConsolidateVCRFilialCreator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.consVCR_filial, header, filialName, false) { }

        protected override void FillReport(CReportVCRFilial[] report, CReportVCRFilial[] yearReport)
        {
            int countReport = report.Length;
            int currentIndex = StartPosition;
            CopyNullCells(ObjWorkSheet, countReport, StartPosition);

            foreach (var data in yearReport)
            {   
                ObjWorkSheet.Cells[currentIndex, 1] = currentIndex-6;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Data._1_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Data._1_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Data._1_total;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Data._11_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Data._11_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Data._11_total;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Data._12_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Data._12_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 11] = data.Data._12_total;
                ObjWorkSheet.Cells[currentIndex, 12] = data.Data._2_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 13] = data.Data._2_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 14] = data.Data._2_total;
                ObjWorkSheet.Cells[currentIndex, 15] = data.Data._21_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 16] = data.Data._21_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 17] = data.Data._21_total;
                ObjWorkSheet.Cells[currentIndex, 18] = data.Data._211_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 19] = data.Data._211_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 20] = data.Data._211_total;
                ObjWorkSheet.Cells[currentIndex, 21] = data.Data._212_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 22] = data.Data._212_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 23] = data.Data._212_total;
                ObjWorkSheet.Cells[currentIndex, 24] = data.Data._213_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 25] = data.Data._213_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 26] = data.Data._213_total;
                ObjWorkSheet.Cells[currentIndex, 27] = data.Data._214_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 28] = data.Data._214_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 29] = data.Data._214_total;
                ObjWorkSheet.Cells[currentIndex, 30] = data.Data._215_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 31] = data.Data._215_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 32] = data.Data._215_total;
                ObjWorkSheet.Cells[currentIndex, 33] = data.Data._216_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 34] = data.Data._216_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 35] = data.Data._216_total;
                ObjWorkSheet.Cells[currentIndex, 36] = data.Data._217_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 37] = data.Data._217_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 38] = data.Data._217_total;
                ObjWorkSheet.Cells[currentIndex, 39] = data.Data._218_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 40] = data.Data._218_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 41] = data.Data._218_total;
                ObjWorkSheet.Cells[currentIndex, 42] = data.Data._219_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 43] = data.Data._219_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 44] = data.Data._219_total;
                ObjWorkSheet.Cells[currentIndex, 45] = data.Data._2110_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 46] = data.Data._2110_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 47] = data.Data._2110_total;
                ObjWorkSheet.Cells[currentIndex, 48] = data.Data._22_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 49] = data.Data._22_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 50] = data.Data._22_total;
                ObjWorkSheet.Cells[currentIndex, 51] = data.Data._221_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 52] = data.Data._221_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 53] = data.Data._221_total;
                ObjWorkSheet.Cells[currentIndex, 54] = data.Data._222_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 55] = data.Data._222_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 56] = data.Data._222_total;
                ObjWorkSheet.Cells[currentIndex, 57] = data.Data._223_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 58] = data.Data._223_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 59] = data.Data._223_total;
                ObjWorkSheet.Cells[currentIndex, 60] = data.Data._224_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 61] = data.Data._224_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 62] = data.Data._224_total;
                ObjWorkSheet.Cells[currentIndex, 63] = data.Data._225_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 64] = data.Data._225_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 65] = data.Data._225_total;
                ObjWorkSheet.Cells[currentIndex, 66] = data.Data._226_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 67] = data.Data._226_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 68] = data.Data._226_total;
                ObjWorkSheet.Cells[currentIndex, 69] = data.Data._227_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 70] = data.Data._227_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 71] = data.Data._227_total;
                ObjWorkSheet.Cells[currentIndex, 72] = data.Data._228_ExpertWithEducation;
                ObjWorkSheet.Cells[currentIndex, 73] = data.Data._228_ExpertWithoutEducation;
                ObjWorkSheet.Cells[currentIndex, 74] = data.Data._228_total;
                currentIndex++;
            }
        }
    }
}
