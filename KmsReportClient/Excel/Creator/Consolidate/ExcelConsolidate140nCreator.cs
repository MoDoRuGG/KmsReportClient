using System;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelConsolidate140nCreator : ExcelBaseCreator<Cons140nTable1[]>
    {
        private const int StartPositionT1 = 12; // Начинаем с 12 строки, т.к. 11-я формула
        private readonly string _period;

        public ExcelConsolidate140nCreator(
            string filename,
            string header,
            string filialName,
            string period) : base(filename, ExcelForm.R140n, header, filialName, false)
        {
            _period = period;
        }

        protected override void FillReport(Cons140nTable1[] report, Cons140nTable1[] yearReport)
        {
            if (ObjWorkSheet != null)
            {
                ObjWorkSheet.Cells[4, 2] = _period;
            }
            // Заполняем Таблицу 1 на первом листе
            FillTable1(report);
        }

        private void FillTable1(Cons140nTable1[] report)
        {
            int currentIndex = StartPositionT1;
            foreach (var data in report)
            {
                // Заполняем данные в строке
                ObjWorkSheet.Cells[currentIndex, 1] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Data.CZLdost;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Data.CZLsmo;

                ObjWorkSheet.Cells[currentIndex, 5] = data.Data.KSErez;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Data.KSE;

                ObjWorkSheet.Cells[currentIndex, 8] = data.Data.PPMinadvn;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Data.Iidvn;

                ObjWorkSheet.Cells[currentIndex, 11] = data.Data.PPMinfdn;
                ObjWorkSheet.Cells[currentIndex, 12] = data.Data.Iidn;

                ObjWorkSheet.Cells[currentIndex, 14] = data.Data.KOJdosud;
                ObjWorkSheet.Cells[currentIndex, 15] = data.Data.KOJsud;
                ObjWorkSheet.Cells[currentIndex, 16] = data.Data.KOJzl;

                ObjWorkSheet.Cells[currentIndex, 18] = data.Data.KOJzlsmo;
                ObjWorkSheet.Cells[currentIndex, 19] = data.Data.CZLsmo;

                ObjWorkSheet.Cells[currentIndex, 21] = data.Data.KZAsobl;
                ObjWorkSheet.Cells[currentIndex, 22] = data.Data.KZAvsego;

                ObjWorkSheet.Cells[currentIndex, 24] = data.Data.DT;
                ObjWorkSheet.Cells[currentIndex, 25] = data.Data.Scpo;

                ObjWorkSheet.Cells[currentIndex, 27] = data.Data.KEKMPpodtv;
                ObjWorkSheet.Cells[currentIndex, 28] = data.Data.KEKMPtfoms;

                ObjWorkSheet.Cells[currentIndex, 30] = data.Data.KZSMOpodtv;
                ObjWorkSheet.Cells[currentIndex, 31] = data.Data.KPMOtfoms;

                currentIndex++;
            }
        }
    }
}