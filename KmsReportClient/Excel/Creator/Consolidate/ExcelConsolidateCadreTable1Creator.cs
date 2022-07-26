using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelConsolidateCadreT1Creator : ExcelBaseCreator<CReportCadreTable1[]>
    {
        private const int StartPosition = 7;

        public ExcelConsolidateCadreT1Creator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.CCadreT1, header, filialName, false) { }

        protected override void FillReport(CReportCadreTable1[] report, CReportCadreTable1[] yearReport)
        {
            int countReport = report.Length;
            int currentIndex = StartPosition;
            CopyNullCells(ObjWorkSheet, countReport, StartPosition);

            foreach (var data in yearReport)
            {
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Data.count_itog_state;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Data.count_itog_fact ;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Data.count_itog_vacancy ;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Data.count_leader_state ;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Data.count_leader_fact ;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Data.count_leader_vacancy ;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Data.count_deputy_leader_state ;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Data.count_deputy_leader_fact ;
                ObjWorkSheet.Cells[currentIndex, 11] = data.Data.count_deputy_leader_vacancy ;
                ObjWorkSheet.Cells[currentIndex, 12] = data.Data.count_expert_doctor_state ;
                ObjWorkSheet.Cells[currentIndex, 13] = data.Data.count_expert_doctor_fact ;
                ObjWorkSheet.Cells[currentIndex, 14] = data.Data.count_expert_doctor_vacancy ;
                ObjWorkSheet.Cells[currentIndex, 15] = data.Data.count_specialist_state ;
                ObjWorkSheet.Cells[currentIndex, 16] = data.Data.count_specialist_fact ;
                ObjWorkSheet.Cells[currentIndex, 17] = data.Data.count_specialist_vacancy ;
                ObjWorkSheet.Cells[currentIndex, 18] = data.Data.count_grf15 ;
                ObjWorkSheet.Cells[currentIndex, 19] = data.Data.count_grf16 ;
                ObjWorkSheet.Cells[currentIndex, 20] = data.Data.count_grf17 ;
                ObjWorkSheet.Cells[currentIndex, 21] = data.Data.count_grf18 ;
                ObjWorkSheet.Cells[currentIndex, 22] = data.Data.count_grf19 ;
                ObjWorkSheet.Cells[currentIndex, 23] = data.Data.count_grf20 ;
                ObjWorkSheet.Cells[currentIndex, 24] = data.Data.count_grf21 ;
                ObjWorkSheet.Cells[currentIndex, 25] = data.Data.count_grf22 ;
                ObjWorkSheet.Cells[currentIndex, 26] = data.Data.count_grf23 ;
                ObjWorkSheet.Cells[currentIndex, 27] = data.Data.count_grf24 ;
                ObjWorkSheet.Cells[currentIndex, 28] = data.Data.count_grf25 ;
                ObjWorkSheet.Cells[currentIndex, 29] = data.Data.count_grf26 ;

            }

        }
    }
}
