using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelFFOMSVolumesByTypesCreator : ExcelBaseCreator<FFOMSVolumesByTypes>
    {
        public ExcelFFOMSVolumesByTypesCreator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.FFOMSVolumesByTypes, header, filialName, false) { }

        protected override void FillReport(FFOMSVolumesByTypes report, FFOMSVolumesByTypes yearReport)
        {
            FillFFOMSVolumesFull(report);
            FillFFOMSVolumesFil(report);
        }

        private void FillFFOMSVolumesFull(FFOMSVolumesByTypes report)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            
            foreach (var data in report.VolFull)
            {
                ObjWorkSheet.Cells[6, 2] = data.mee_unpl;
                ObjWorkSheet.Cells[7, 2] = data.mee_pl;
                ObjWorkSheet.Cells[10, 2] = data.ekmp_unpl;
                ObjWorkSheet.Cells[11, 2] = data.ekmp_pl;

                ObjWorkSheet.Cells[5, 4] = data.mek_app;
                ObjWorkSheet.Cells[5, 5] = data.mee_app_unpl;
                ObjWorkSheet.Cells[5, 6] = data.mee_app_pl;

                ObjWorkSheet.Cells[6, 4] = data.mek_skp;
                ObjWorkSheet.Cells[6, 5] = data.mee_skp_unpl;
                ObjWorkSheet.Cells[6, 6] = data.mee_skp_pl;

                ObjWorkSheet.Cells[7, 4] = data.mek_sdp;
                ObjWorkSheet.Cells[7, 5] = data.mee_sdp_unpl;
                ObjWorkSheet.Cells[7, 6] = data.mee_sdp_pl;

                ObjWorkSheet.Cells[8, 4] = data.mek_smp;
                ObjWorkSheet.Cells[8, 5] = data.mee_smp_unpl;
                ObjWorkSheet.Cells[8, 6] = data.mee_smp_pl;


                ObjWorkSheet.Cells[9, 4] = data.mek_app;
                ObjWorkSheet.Cells[9, 5] = data.ekmp_app_unpl;
                ObjWorkSheet.Cells[9, 6] = data.ekmp_app_pl;

                ObjWorkSheet.Cells[10, 4] = data.mek_skp;
                ObjWorkSheet.Cells[10, 5] = data.ekmp_skp_unpl;
                ObjWorkSheet.Cells[10, 6] = data.ekmp_skp_pl;

                ObjWorkSheet.Cells[11, 4] = data.mek_sdp;
                ObjWorkSheet.Cells[11, 5] = data.ekmp_sdp_unpl;
                ObjWorkSheet.Cells[11, 6] = data.ekmp_sdp_pl;

                ObjWorkSheet.Cells[12, 4] = data.mek_smp;
                ObjWorkSheet.Cells[12, 5] = data.ekmp_smp_unpl;
                ObjWorkSheet.Cells[12, 6] = data.ekmp_smp_pl;
            }
        }


        private void FillFFOMSVolumesFil(FFOMSVolumesByTypes report)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[2];
            int countReport = report.VolFil.Length;
            int currentIndex = 3;

            foreach (var data in report.VolFil.OrderBy(x => x.Filial))
            {
                if (data.Filial != "")
                {
                    ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                    ObjWorkSheet.Cells[currentIndex, 3] = data.mek_app;
                    ObjWorkSheet.Cells[currentIndex, 4] = data.mee_app;
                    ObjWorkSheet.Cells[currentIndex, 6] = data.ekmp_app;

                    ObjWorkSheet.Cells[currentIndex, 8] = data.mek_skp;
                    ObjWorkSheet.Cells[currentIndex, 9] = data.mee_skp;
                    ObjWorkSheet.Cells[currentIndex, 11] = data.ekmp_skp;

                    ObjWorkSheet.Cells[currentIndex, 13] = data.mek_smp;
                    ObjWorkSheet.Cells[currentIndex, 14] = data.mee_smp;
                    ObjWorkSheet.Cells[currentIndex, 16] = data.ekmp_smp;

                    ObjWorkSheet.Cells[currentIndex, 18] = data.mek_sdp;
                    ObjWorkSheet.Cells[currentIndex, 19] = data.mee_sdp;
                    ObjWorkSheet.Cells[currentIndex, 21] = data.ekmp_sdp;

                    currentIndex++;
                }
            }
        }
    }
}
