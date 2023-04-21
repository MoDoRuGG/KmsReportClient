using KmsReportClient.External;
using KmsReportClient.Model.Enums;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelConsZpzWebSite2023 : ExcelBaseCreator<ZpzForWebSite2023>
    {
        public ExcelConsZpzWebSite2023(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.ZpzForWebSite2023, header, filialName, true) { }

        protected override void FillReport(ZpzForWebSite2023 report, ZpzForWebSite2023 yearReport)
        {
            ObjWorkSheet.Cells[2, 2] = FilialName;

            FillTreatments(report);
            FillComplaints(report);
            FillProtections(report);
            FillExpertises(report);
            FillSpecialists(report);
            FillInformations(report);
        }

        private void FillTreatments(ZpzForWebSite2023 report)
        {
            int currentIndex = 7;
            foreach (var treatment in report.Treatments)
            {
                ObjWorkSheet.Cells[++currentIndex, 3] = treatment.Oral;
                ObjWorkSheet.Cells[currentIndex, 4] = treatment.Written;
            }
        }

        private void FillComplaints(ZpzForWebSite2023 report)
        {
            int currentIndex = 15;
            foreach (var complaint in report.Complaints)
            {
                ObjWorkSheet.Cells[++currentIndex, 3] = complaint.Oral;
                ObjWorkSheet.Cells[currentIndex, 4] = complaint.Written;
            }
        }

        private void FillExpertises(ZpzForWebSite2023 report)
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

        private void FillProtections(ZpzForWebSite2023 report)
        {
            int currentIndex = 44;
            foreach (var protection in report.Protections)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = protection.Count;
            }
        }

        private void FillSpecialists(ZpzForWebSite2023 report)
        {
            int currentIndex = 49;
            foreach (var specialist in report.Specialists)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = specialist.Count;
            }
        }

        private void FillInformations(ZpzForWebSite2023 report)
        {
            int currentIndex = 57;
            foreach (var information in report.Informations)
            {
                ObjWorkSheet.Cells[currentIndex++, 3] = information.Count;
            }
        }
    }
}
