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
    class ExcelConsolidateCardio : ExcelBaseCreator<ConsolidateCardio[]>
    {
        private const int StartPosition = 5;

        public ExcelConsolidateCardio(
           string filename,
           string header,
           string filialName) : base(filename, ExcelForm.cardio, header, filialName, false) { }

        protected override void FillReport(ConsolidateCardio[] report, ConsolidateCardio[] yearReport)
        {
            FillEmkpMee(report);
            FillComplaint(report);
            FillProtection(report);
            FillFinance(report);
            
        }

        private void FillComplaint(ConsolidateCardio[] reports)
        {
            var complaint = reports.Select(r => new { r.Filial, r.Complaint }).ToList();

            int countReport = complaint.Count;
            int currentIndex = StartPosition;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            CopyNullCells(ObjWorkSheet, countReport, StartPosition);

            int counter = 1;
            foreach (var data in complaint)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Complaint.MedicalHelp;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Complaint.Underage;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Complaint.AskMedicalHelp;
                ObjWorkSheet.Cells[currentIndex++, 6] = data.Complaint.UnderageAskMedicalHelp;

            }

        }

        private void FillProtection(ConsolidateCardio[] reports)
        {
            var protection = reports.Select(r => new { r.Filial, r.Protection }).ToList();

            int countReport = protection.Count;
            int currentIndex = StartPosition;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[2];
            CopyNullCells(ObjWorkSheet, countReport, StartPosition);

            int counter = 1;
            foreach (var data in protection)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Protection.PretrialMedicalHelp;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Protection.PretrialUnderage;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Protection.JudicalMedicalHelp;
                ObjWorkSheet.Cells[currentIndex++, 6] = data.Protection.JudicalUnderage;

            }
        }

        private void FillFinance(ConsolidateCardio[] reports)
        {
            var finance = reports.Select(r => new { r.Filial, r.Finance }).ToList();

            int countReport = finance.Count;
            int currentIndex = StartPosition;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[4];
            CopyNullCells(ObjWorkSheet, countReport, StartPosition);

            int counter = 1;
            foreach (var data in finance)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Finance.SumMeeNotTimeDispan;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Finance.SumMeeNotTimeDispanUnderage;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Finance.SumEkmpNotTimeDispan;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Finance.SumEkmpNotTimeDispanUnderage;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Finance.SumNeprofilHelp;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Finance.SumNeprofilHelpUnderage;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Finance.SumNevipolnenie;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Finance.SumNevipolnenieUnderage;
                ObjWorkSheet.Cells[currentIndex, 11] = data.Finance.SumNesobludRecomedation;
                ObjWorkSheet.Cells[currentIndex, 12] = data.Finance.SumNesobludRecomedationUnderage;
                ObjWorkSheet.Cells[currentIndex, 13] = data.Finance.SumCloseHelp;
                ObjWorkSheet.Cells[currentIndex, 14] = data.Finance.SumCloseHelpUnderage;
                ObjWorkSheet.Cells[currentIndex, 15] = data.Finance.SumViolation;
                ObjWorkSheet.Cells[currentIndex++, 16] = data.Finance.SumViolationUnderage;


            }
        }


        private void FillEmkpMee(ConsolidateCardio[] reports)
        {
            var ekmpMee = reports.Select(r => new { r.Filial, r.MeeEkmp }).ToList();

            int countReport = ekmpMee.Count;
            int currentIndex = StartPosition;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[3];
            CopyNullCells(ObjWorkSheet, countReport, StartPosition);
          

            int counter = 1;

            foreach (var data in ekmpMee)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.MeeEkmp.ComplaintsMEE;
                ObjWorkSheet.Cells[currentIndex, 4] = data.MeeEkmp.HospitalizationMEE;
                ObjWorkSheet.Cells[currentIndex, 5] = data.MeeEkmp.CelExpertiseMEE;
                ObjWorkSheet.Cells[currentIndex, 6] = data.MeeEkmp.PlanExpertiseMEE;
                ObjWorkSheet.Cells[currentIndex, 7] = data.MeeEkmp.ComplaintsEKMP;
                ObjWorkSheet.Cells[currentIndex, 8] = data.MeeEkmp.CelLetalKoronar;
                ObjWorkSheet.Cells[currentIndex, 9] = data.MeeEkmp.CelLetalOnmk;
                ObjWorkSheet.Cells[currentIndex, 10] = data.MeeEkmp.CelAllCardio;
                ObjWorkSheet.Cells[currentIndex, 11] = data.MeeEkmp.CelUnderage;
                ObjWorkSheet.Cells[currentIndex, 12] = data.MeeEkmp.PlanAllCardio;
                ObjWorkSheet.Cells[currentIndex, 13] = data.MeeEkmp.PlanCardioUnderage;
                ObjWorkSheet.Cells[currentIndex, 14] = data.MeeEkmp.CelNeprofilGospital;
                ObjWorkSheet.Cells[currentIndex, 15] = data.MeeEkmp.CelNeprofilGospitalUnderage;
                ObjWorkSheet.Cells[currentIndex, 16] = data.MeeEkmp.PlanNeprofilGospital;
                ObjWorkSheet.Cells[currentIndex, 17] = data.MeeEkmp.PlanNeprofilGospitalUnderage;
                ObjWorkSheet.Cells[currentIndex, 18] = data.MeeEkmp.CelNevipolnenie;
                ObjWorkSheet.Cells[currentIndex, 19] = data.MeeEkmp.CelNevipolnenieUnderage;
                ObjWorkSheet.Cells[currentIndex, 20] = data.MeeEkmp.PlanNevipolnenie;
                ObjWorkSheet.Cells[currentIndex, 21] = data.MeeEkmp.PlanNevipolnenieUnderage;
                ObjWorkSheet.Cells[currentIndex, 22] = data.MeeEkmp.CelNotAddDispNab;
                ObjWorkSheet.Cells[currentIndex, 23] = data.MeeEkmp.CelNotAddDispNabUnderage;
                ObjWorkSheet.Cells[currentIndex, 24] = data.MeeEkmp.PlanNotAddDispNab;
                ObjWorkSheet.Cells[currentIndex, 25] = data.MeeEkmp.PlanNotAddDispNabUnderage;
                ObjWorkSheet.Cells[currentIndex, 26] = data.MeeEkmp.CelNotSobludClinicRecomendation;
                ObjWorkSheet.Cells[currentIndex, 27] = data.MeeEkmp.CelNotSobludClinicRecomendationUnderage;
                ObjWorkSheet.Cells[currentIndex, 28] = data.MeeEkmp.PlanNotSobludClinicRecomendation;
                ObjWorkSheet.Cells[currentIndex, 29] = data.MeeEkmp.PlanNotSobludClinicRecomendationUnderage;
                ObjWorkSheet.Cells[currentIndex, 30] = data.MeeEkmp.CelPrematureCloseHelpMerop;
                ObjWorkSheet.Cells[currentIndex, 31] = data.MeeEkmp.CelPrematureCloseHelpMeropUnderage;
                ObjWorkSheet.Cells[currentIndex, 32] = data.MeeEkmp.PlanPrematureCloseHelpMerop;
                ObjWorkSheet.Cells[currentIndex, 33] = data.MeeEkmp.PlanPrematureCloseHelpMeropUnderage;
                ObjWorkSheet.Cells[currentIndex, 34] = data.MeeEkmp.CelViolationHospital;
                ObjWorkSheet.Cells[currentIndex, 35] = data.MeeEkmp.CelViolationHospitalUnderage;
                ObjWorkSheet.Cells[currentIndex, 36] = data.MeeEkmp.PlanViolationHospital;
                ObjWorkSheet.Cells[currentIndex, 37] = data.MeeEkmp.PlanViolationHospitalUnderage;
                ObjWorkSheet.Cells[currentIndex, 38] = data.MeeEkmp.CelNeobosOtkaz;
                ObjWorkSheet.Cells[currentIndex, 39] = data.MeeEkmp.CelNeobosOtkazUnderage;
                ObjWorkSheet.Cells[currentIndex, 40] = data.MeeEkmp.PlanNeobosOtkaz;
                ObjWorkSheet.Cells[currentIndex++, 41] = data.MeeEkmp.PlanNeobosOtkazUnderage;


            }

        }



    }
}
