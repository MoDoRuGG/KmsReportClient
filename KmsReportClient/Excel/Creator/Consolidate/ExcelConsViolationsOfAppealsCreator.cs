using KmsReportClient.External;
using KmsReportClient.Model.Enums;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelConsViolationsOfAppealsCreator : ExcelBaseCreator<ViolationsOfAppeals>
    {
        public ExcelConsViolationsOfAppealsCreator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.ZpzForWebSite2023, header, filialName, true) { }

        protected override void FillReport(ViolationsOfAppeals report, ViolationsOfAppeals yearReport)
        {
            ObjWorkSheet.Cells[2, 2] = FilialName;

            FillTreatments(report);
            FillComplaints(report);
            FillProtections(report);
            FillExpertises(report);
            FillSpecialists(report);
            FillInformations(report);
        }

        private void FillTreatments(ViolationsOfAppeals report)
        {
            int currentIndex = 7;
            foreach (var treatment in report.Treatments)
            {
                ObjWorkSheet.Cells[currentIndex, 3] = treatment.Oral;
                ObjWorkSheet.Cells[currentIndex, 4] = treatment.Written;
                ObjWorkSheet.Cells[currentIndex++, 5] = treatment.Assignment;
            }
        }

        private void FillComplaints(ViolationsOfAppeals report)
        {
            int currentIndex = 12;
            foreach (var complaint in report.Complaints)
            {
                ObjWorkSheet.Cells[currentIndex, 3] = complaint.Oral;
                ObjWorkSheet.Cells[currentIndex, 4] = complaint.Written;
                ObjWorkSheet.Cells[currentIndex++, 5] = complaint.Assignment;
            }
        }

        private void FillExpertises(ViolationsOfAppeals report)
        {
            int currentIndex = 28;
            foreach (var expertise in report.Expertises)
            {
                ObjWorkSheet.Cells[++currentIndex, 3] = expertise.Target;
                ObjWorkSheet.Cells[++currentIndex, 3] = expertise.Plan;
                ObjWorkSheet.Cells[++currentIndex, 3] = expertise.Violation;
                currentIndex += 2;
            }
        }

        private void FillProtections(ViolationsOfAppeals report)
        {
            int currentIndex = 41;
            foreach (var protection in report.Protections)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = protection.Count;
            }
        }

        private void FillSpecialists(ViolationsOfAppeals report)
        {
            int currentIndex = 46;
            foreach (var specialist in report.Specialists)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = specialist.Count;
            }
        }

        private void FillInformations(ViolationsOfAppeals report)
        {
            int currentIndex = 54;
            foreach (var information in report.Informations)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = information.Count;
            }
        }
    }
}
