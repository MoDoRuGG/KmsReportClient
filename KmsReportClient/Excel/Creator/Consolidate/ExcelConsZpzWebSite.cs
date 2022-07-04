using KmsReportClient.External;
using KmsReportClient.Model.Enums;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelConsZpzWebSite : ExcelBaseCreator<ZpzForWebSite>
    {
        public ExcelConsZpzWebSite(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.ZpzForWebSite, header, filialName, true) { }

        protected override void FillReport(ZpzForWebSite report, ZpzForWebSite yearReport)
        {
            ObjWorkSheet.Cells[2, 2] = FilialName;

            FillTreatments(report);
            FillComplaints(report);
            FillProtections(report);
            FillExpertises(report);
            FillSpecialists(report);
            FillComplacence(report);
            FillInformations(report);
        }

        private void FillTreatments(ZpzForWebSite report)
        {
            int currentIndex = 7;
            foreach (var treatment in report.Treatments)
            {
                ObjWorkSheet.Cells[++currentIndex, 3] = treatment.Oral;
                ObjWorkSheet.Cells[currentIndex, 4] = treatment.Written;
            }
        }

        private void FillComplaints(ZpzForWebSite report)
        {
            int currentIndex = 15;
            foreach (var complaint in report.Complaints)
            {
                ObjWorkSheet.Cells[++currentIndex, 3] = complaint.Oral;
                ObjWorkSheet.Cells[currentIndex, 4] = complaint.Written;
            }
        }

        private void FillExpertises(ZpzForWebSite report)
        {
            int currentIndex = 31;
            foreach (var expertise in report.Expertises)
            {
                ObjWorkSheet.Cells[++currentIndex, 3] = expertise.Target;
                ObjWorkSheet.Cells[++currentIndex, 3] = expertise.Plan;
                ObjWorkSheet.Cells[++currentIndex, 3] = expertise.Violation;
                currentIndex += 2;
            }
        }

        private void FillProtections(ZpzForWebSite report)
        {
            int currentIndex = 44;
            foreach (var protection in report.Protections)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = protection.Count;
            }
        }

        private void FillSpecialists(ZpzForWebSite report)
        {
            int currentIndex = 49;
            foreach (var specialist in report.Specialists)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = specialist.Count;
            }
        }

        private void FillComplacence(ZpzForWebSite report)
        {
            int currentIndex = 57;
            foreach (var complacence in report.Complacence)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = complacence.Count;
            }
        }

        private void FillInformations(ZpzForWebSite report)
        {
            int currentIndex = 67;
            foreach (var information in report.Informations)
            {
                ObjWorkSheet.Cells[++currentIndex, 3] = information.Count;
            }
        }
    }
}
