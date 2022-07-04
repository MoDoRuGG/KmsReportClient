using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelOnkoCreator : ExcelBaseCreator<ConsolidateOnko[]>
    {
        private const int StartPosition = 6;
       
        public ExcelOnkoCreator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.Onko, header, filialName, false) { }

        protected override void FillReport(ConsolidateOnko[] report, ConsolidateOnko[] yearReport)
        {
            FillComplaint(report);
            FillProtection(report);
            FillMek(report);
            FillMee(report);
            FillEkmp(report);
            FillFinance(report);
        }

        private void FillComplaint(ConsolidateOnko[] reports)
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
                ObjWorkSheet.Cells[currentIndex, 4] = data.Complaint.Dedlines;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Complaint.MedicineProvision;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Complaint.DedlineDrugMedicine;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Complaint.NoDrugMedicine;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Complaint.AppealMedicalHelp;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Complaint.AppealMedicineProvision;
                ObjWorkSheet.Cells[currentIndex++, 10] = data.Complaint.AppealDrugsMedicine;
            }
        }

        private void FillProtection(ConsolidateOnko[] reports)
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
                ObjWorkSheet.Cells[currentIndex, 4] = data.Protection.PretrialDeadline;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Protection.PretrialMedicineProvision;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Protection.PretrialDedlineDrugMedicine;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Protection.PretrialNoDrugMedicine;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Protection.JudicalMedicalHelp;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Protection.JudicalDeadline;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Protection.JudicalMedicineProvision;
                ObjWorkSheet.Cells[currentIndex, 11] = data.Protection.JudicalDedlineDrugMedicine;
                ObjWorkSheet.Cells[currentIndex++, 12] = data.Protection.JudicalNoDrugMedicine;
            }
        }

        private void FillMek(ConsolidateOnko[] reports)
        {
            var mek = reports.Select(r => new { r.Filial, r.Mek }).ToList();

            int countReport = mek.Count;
            int currentIndex = StartPosition;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[3];
            CopyNullCells(ObjWorkSheet, countReport, StartPosition);

            int counter = 1;
            foreach (var data in mek)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Mek.PresentedBills;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Mek.AcceptedBills;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Mek.RegistrationMek;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Mek.NotInProgramMek;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Mek.TarifMek;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Mek.LicenceMek;
                ObjWorkSheet.Cells[currentIndex++, 9] = data.Mek.RepeatMek;
            }
        }

        private void FillMee(ConsolidateOnko[] reports)
        {
            var mee = reports.Select(r => new { r.Filial, r.Mee }).ToList();

            int countReport = mee.Count;
            int currentIndex = StartPosition;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[4];
            CopyNullCells(ObjWorkSheet, countReport, StartPosition);

            int counter = 1;
            foreach (var data in mee)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Mee.Complaint;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Mee.Antitumor;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Mee.PlanHosp;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Mee.ViolationCondition;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Mee.ViolationOnkoFirst;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Mee.ViolationHisto;
                ObjWorkSheet.Cells[currentIndex++, 9] = data.Mee.ViolationOnkoDiagnostic;
            }
        }

        private void FillEkmp(ConsolidateOnko[] reports)
        {
            var ekmp = reports.Select(r => new { r.Filial, r.Ekmp }).ToList();

            int countReport = ekmp.Count;
            int currentIndex = StartPosition;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[5];
            CopyNullCells(ObjWorkSheet, countReport, StartPosition);

            int counter = 1;
            foreach (var data in ekmp)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Ekmp.EkmpComplaint;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Ekmp.EkmpFromMee;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Ekmp.DeathEkmp;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Ekmp.ThematicEkmp;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Ekmp.CountOnko;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Ekmp.NoProfilOnko;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Ekmp.UnreasonEkmp;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Ekmp.DispEkmp;
                ObjWorkSheet.Cells[currentIndex, 11] = data.Ekmp.RecommendationEkmp;
                ObjWorkSheet.Cells[currentIndex, 12] = data.Ekmp.Premature;
                ObjWorkSheet.Cells[currentIndex, 13] = data.Ekmp.ViolationMoEmkp;
                ObjWorkSheet.Cells[currentIndex, 14] = data.Ekmp.Failure;
                ObjWorkSheet.Cells[currentIndex, 15] = data.Ekmp.Payment;
                ObjWorkSheet.Cells[currentIndex++, 16] = data.Ekmp.OtherViolation;
            }
        }

        private void FillFinance(ConsolidateOnko[] reports)
        {
            var finance = reports.Select(r => new { r.Filial, r.Finance }).ToList();

            int countReport = finance.Count;
            int currentIndex = StartPosition;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[6];
            CopyNullCells(ObjWorkSheet, countReport, StartPosition);

            int counter = 1;
            foreach (var data in finance)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Finance.SumMek;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Finance.SumDispMee;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Finance.SumMee;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Finance.SumDispEkmp;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Finance.SumNoProfilEkmp;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Finance.SumUnreasonEkmp;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Finance.SumRecommendationEkmp;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Finance.SumPrematureEkmp;
                ObjWorkSheet.Cells[currentIndex, 11] = data.Finance.SumFailureEkmp;
                ObjWorkSheet.Cells[currentIndex, 12] = data.Finance.SumPaymentEkmp;
                ObjWorkSheet.Cells[currentIndex++, 13] = data.Finance.SumOtherEkmp;
            }
        }


    }
}
