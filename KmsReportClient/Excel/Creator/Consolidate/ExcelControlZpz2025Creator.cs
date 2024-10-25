using System;
using System.Collections.Generic;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelControlZpz2025Creator : ExcelBaseCreator<CReportZpz2025[]>
    {
        private const int StartPosition = 6;
        private const int StartPersonnel = 4;
        private string _yymm;

        public ExcelControlZpz2025Creator(
            string filename,
            string header,
            string filialName, string yymm) : base(filename, ExcelForm.ControlZpz2025, header, filialName, false) { _yymm = yymm; }

        protected override void FillReport(CReportZpz2025[] reports, CReportZpz2025[] yearReport)
        {
            FillExpertises(reports);
            FillFinance(reports);
            FillPersonnel(reports);
            FillNormative(reports);
        }

        private void FillNormative(CReportZpz2025[] reports)
        {
            Dictionary<int, double> newNormative = new Dictionary<int, double>()
            {
                { 6, 2},
                { 10, 0.5},
                { 16, 6},
                { 22, 6},
                { 25, 0.5},
                { 28, 0.2},
                { 33, 1.5},
                { 38, 3},
            };

            var normative = reports.Select(r => new { r.Filial, r.Normative }).ToList();

            int countReport = normative.Count;
            int currentIndex = StartPosition;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[2];

            CopyNullCells(ObjWorkSheet, countReport, StartPosition);

            foreach (var item in newNormative) {ObjWorkSheet.Cells[2, item.Key] = item.Value;}

            int counter = 1;
            foreach (var data in normative)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Normative.BillsOutMo;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Normative.MeeOutMoPlan;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Normative.MeeOutMoTarget;
                
                ObjWorkSheet.Cells[currentIndex, 7] = data.Normative.BillsApp;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Normative.MeeAppPlan;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Normative.MeeAppTarget;
                
                ObjWorkSheet.Cells[currentIndex, 11] = data.Normative.BillsDayHosp;
                ObjWorkSheet.Cells[currentIndex, 12] = data.Normative.MeeDayHospTarget;
                ObjWorkSheet.Cells[currentIndex, 13] = data.Normative.MeeDayHospTargetVmp;
                ObjWorkSheet.Cells[currentIndex, 14] = data.Normative.MeeDayHospPlan;
                ObjWorkSheet.Cells[currentIndex, 15] = data.Normative.MeeDayHospPlanVmp;
                
                ObjWorkSheet.Cells[currentIndex, 17] = data.Normative.BillsHosp;
                ObjWorkSheet.Cells[currentIndex, 18] = data.Normative.MeeHospTarget;
                ObjWorkSheet.Cells[currentIndex, 19] = data.Normative.MeeHospTargetVmp;
                ObjWorkSheet.Cells[currentIndex, 20] = data.Normative.MeeHospPlan;
                ObjWorkSheet.Cells[currentIndex, 21] = data.Normative.MeeHospPlanVmp;

                ObjWorkSheet.Cells[currentIndex, 23] = data.Normative.EkmpOutMoPlan;
                ObjWorkSheet.Cells[currentIndex, 24] = data.Normative.EkmpOutMoTarget;
                
                ObjWorkSheet.Cells[currentIndex, 26] = data.Normative.EkmpAppPlan;
                ObjWorkSheet.Cells[currentIndex, 27] = data.Normative.EkmpAppTarget;

                ObjWorkSheet.Cells[currentIndex, 29] = data.Normative.EkmpDayHospTarget;
                ObjWorkSheet.Cells[currentIndex, 30] = data.Normative.EkmpDayHospTargetVmp;
                ObjWorkSheet.Cells[currentIndex, 31] = data.Normative.EkmpDayHospPlan;
                ObjWorkSheet.Cells[currentIndex, 32] = data.Normative.EkmpDayHospPlanVmp;

                ObjWorkSheet.Cells[currentIndex, 34] = data.Normative.EkmpHospTarget;
                ObjWorkSheet.Cells[currentIndex, 35] = data.Normative.EkmpHospTargetVmp;
                ObjWorkSheet.Cells[currentIndex, 36] = data.Normative.EkmpHospPlan;
                ObjWorkSheet.Cells[currentIndex++, 37] = data.Normative.EkmpHospPlanVmp;
            }
        }

        private void FillPersonnel(CReportZpz2025[] reports)
        {
            var personnel = reports.Select(r => new { r.Filial, r.Personnel }).ToList();

            int countReport = personnel.Count;
            int currentIndex = StartPersonnel;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[4];
            CopyNullCells(ObjWorkSheet, countReport, StartPersonnel);

            int counter = 1;
            foreach (var data in personnel)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Personnel.SpecialistFullTime;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Personnel.SpecialistRemote;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Personnel.FullTime;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Personnel.Remote;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Personnel.ExpertsFullTime;
                ObjWorkSheet.Cells[currentIndex, 11] = data.Personnel.ExpertsRemote;
                ObjWorkSheet.Cells[currentIndex, 13] = data.Personnel.ExpertsEkmpRegion;
                ObjWorkSheet.Cells[currentIndex, 14] = data.Personnel.ExpertsEkmpRemote;
                ObjWorkSheet.Cells[currentIndex, 16] = data.Personnel.ExpertisesFullTime;
                ObjWorkSheet.Cells[currentIndex, 17] = data.Personnel.ExpertisesRemote;
                ObjWorkSheet.Cells[currentIndex, 19] = data.Personnel.ExpertisesPlanFullTime;
                ObjWorkSheet.Cells[currentIndex, 20] = data.Personnel.ExpertisesPlanRemote;
                ObjWorkSheet.Cells[currentIndex, 22] = data.Personnel.ExpertisesPlanAppealFullTime;
                ObjWorkSheet.Cells[currentIndex, 23] = data.Personnel.ExpertisesPlanAppealRemote;
                ObjWorkSheet.Cells[currentIndex, 25] = data.Personnel.ExpertisesPlanUnfoundedFullTime;
                ObjWorkSheet.Cells[currentIndex, 26] = data.Personnel.ExpertisesPlanUnfoundedRemote;
                ObjWorkSheet.Cells[currentIndex, 28] = data.Personnel.ExpertisesUnplannedFullTime;
                ObjWorkSheet.Cells[currentIndex, 29] = data.Personnel.ExpertisesUnplannedRemote;
                ObjWorkSheet.Cells[currentIndex, 31] = data.Personnel.ExpertisesUnplannedAppealFullTime;
                ObjWorkSheet.Cells[currentIndex, 32] = data.Personnel.ExpertisesUnplannedAppealRemote;
                ObjWorkSheet.Cells[currentIndex, 34] = data.Personnel.ExpertisesUnplannedUnfoundedFullTime;
                ObjWorkSheet.Cells[currentIndex, 35] = data.Personnel.ExpertisesUnplannedUnfoundedRemote;
                ObjWorkSheet.Cells[currentIndex, 37] = data.Personnel.ExpertisesThemeFullTime;
                ObjWorkSheet.Cells[currentIndex, 38] = data.Personnel.ExpertisesThemeRemote;
                ObjWorkSheet.Cells[currentIndex, 40] = data.Personnel.ExpertisesThemeAppealFullTime;
                ObjWorkSheet.Cells[currentIndex, 41] = data.Personnel.ExpertisesThemeAppealRemote;
                ObjWorkSheet.Cells[currentIndex, 43] = data.Personnel.ExpertisesThemeUnfoundedFullTime;
                ObjWorkSheet.Cells[currentIndex, 44] = data.Personnel.ExpertisesThemeUnfoundedRemote;
                ObjWorkSheet.Cells[currentIndex, 46] = data.Personnel.ExpertisesMultiFullTime;
                ObjWorkSheet.Cells[currentIndex, 47] = data.Personnel.ExpertisesMultiRemote;
                ObjWorkSheet.Cells[currentIndex, 49] = data.Personnel.ExpertisesMultiAppealFullTime;
                ObjWorkSheet.Cells[currentIndex, 50] = data.Personnel.ExpertisesMultiAppealRemote;
                ObjWorkSheet.Cells[currentIndex, 52] = data.Personnel.ExpertisesMultiUnfoundedFullTime;
                ObjWorkSheet.Cells[currentIndex, 53] = data.Personnel.ExpertisesMultiUnfoundedRemote;
                ObjWorkSheet.Cells[currentIndex, 55] = data.Personnel.PreparedFullTime;
                ObjWorkSheet.Cells[currentIndex, 56] = data.Personnel.PreparedRemote;
                ObjWorkSheet.Cells[currentIndex, 58] = data.Personnel.QualFullTime;
                ObjWorkSheet.Cells[currentIndex, 59] = data.Personnel.QualRemote;
                ObjWorkSheet.Cells[currentIndex, 61] = data.Personnel.QualHigherFullTime;
                ObjWorkSheet.Cells[currentIndex, 62] = data.Personnel.QualHigherRemote;
                ObjWorkSheet.Cells[currentIndex, 64] = data.Personnel.Qual1stFullTime;
                ObjWorkSheet.Cells[currentIndex, 65] = data.Personnel.Qual1stRemote;
                ObjWorkSheet.Cells[currentIndex, 67] = data.Personnel.Qual2ndFullTime;
                ObjWorkSheet.Cells[currentIndex, 68] = data.Personnel.Qual2ndRemote;
                ObjWorkSheet.Cells[currentIndex, 70] = data.Personnel.DegreeFullTime;
                ObjWorkSheet.Cells[currentIndex, 71] = data.Personnel.DegreeRemote;
                ObjWorkSheet.Cells[currentIndex, 73] = data.Personnel.CandidateFullTime;
                ObjWorkSheet.Cells[currentIndex, 74] = data.Personnel.CandidateRemote;
                ObjWorkSheet.Cells[currentIndex, 76] = data.Personnel.DoctorFullTime;
                ObjWorkSheet.Cells[currentIndex, 77] = data.Personnel.DoctorRemote;
                ObjWorkSheet.Cells[currentIndex, 79] = data.Personnel.InsRepresFullTime;
                ObjWorkSheet.Cells[currentIndex, 82] = data.Personnel.InsRepres1FullTime;
                ObjWorkSheet.Cells[currentIndex, 85] = data.Personnel.InsRepres1spFullTime;
                ObjWorkSheet.Cells[currentIndex, 88] = data.Personnel.InsRepres2FullTime;
                ObjWorkSheet.Cells[currentIndex, 91] = data.Personnel.InsRepres2spFullTime;
                ObjWorkSheet.Cells[currentIndex, 94] = data.Personnel.InsRepres3FullTime;
                ObjWorkSheet.Cells[currentIndex++, 97] = data.Personnel.InsRepres3spFullTime;

            }
        }

        private void FillFinance(CReportZpz2025[] reports)
        {
            var finance = reports.Select(r => new { r.Filial, r.Finance }).ToList();

            int countReport = finance.Count;
            int currentIndex = StartPosition;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[3];
            CopyNullCells(ObjWorkSheet, countReport, StartPosition);

            int counter = 1;
            foreach (var data in finance)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Finance.SumPayment;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Finance.SumNotPayment;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Finance.SumMek;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Finance.SumMee;
                ObjWorkSheet.Cells[currentIndex++, 7] = data.Finance.SumEkmp;
            }
        }

        private void FillExpertises(CReportZpz2025[] reports)
        {
            var expertises = reports.Select(r => new { r.Filial, r.Expertise }).ToList();

            int countReport = expertises.Count;
            int currentIndex = StartPosition;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            CopyNullCells(ObjWorkSheet, countReport, StartPosition);

            foreach (var data in expertises) 
            {
                ObjWorkSheet.Cells[currentIndex, 1] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Expertise.Bills;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Expertise.CountMeeTarget;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Expertise.CountMeePlan;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Expertise.CountMeeComplaintTarget;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Expertise.CountMeeComplaintPlan;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Expertise.CountMeeRepeat;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Expertise.CountMeeOnco;
                ObjWorkSheet.Cells[currentIndex, 11] = data.Expertise.CountMeeDs;
                ObjWorkSheet.Cells[currentIndex, 12] = data.Expertise.CountMeeLeth;
                ObjWorkSheet.Cells[currentIndex, 13] = data.Expertise.CountMeeInjured;
                ObjWorkSheet.Cells[currentIndex, 14] = data.Expertise.CountMeeDefectedCaseTarget;
                ObjWorkSheet.Cells[currentIndex, 15] = data.Expertise.CountMeeDefectedCasePlan;
                ObjWorkSheet.Cells[currentIndex, 17] = data.Expertise.CountMeeDefectsTarget;
                ObjWorkSheet.Cells[currentIndex, 18] = data.Expertise.CountMeeDefectsPlan;
                ObjWorkSheet.Cells[currentIndex, 20] = data.Expertise.CountMeeDefectsPeriod;
                ObjWorkSheet.Cells[currentIndex, 21] = data.Expertise.CountMeeDefectsCondition;
                ObjWorkSheet.Cells[currentIndex, 22] = data.Expertise.CountMeeDefectsRepeat;
                ObjWorkSheet.Cells[currentIndex, 23] = data.Expertise.CountMeeDefectsOutOfDocums;
                ObjWorkSheet.Cells[currentIndex, 24] = data.Expertise.CountMeeDefectsUnpayable;
                ObjWorkSheet.Cells[currentIndex, 25] = data.Expertise.CountMeeDefectsBuyMedicament;
                ObjWorkSheet.Cells[currentIndex, 26] = data.Expertise.CountMeeDefectsOutOfLeth;
                ObjWorkSheet.Cells[currentIndex, 27] = data.Expertise.CountMeeDefectsWithoutDocums;
                ObjWorkSheet.Cells[currentIndex, 28] = data.Expertise.CountMeeDefectsIncorrectDocums;
                ObjWorkSheet.Cells[currentIndex, 29] = data.Expertise.CountMeeDefectsBadDocums;
                ObjWorkSheet.Cells[currentIndex, 30] = data.Expertise.CountMeeDefectsBadDate;
                ObjWorkSheet.Cells[currentIndex, 31] = data.Expertise.CountMeeDefectsBadData;
                ObjWorkSheet.Cells[currentIndex, 32] = data.Expertise.CountMeeDefectsOutOfProtocol;
                ObjWorkSheet.Cells[currentIndex, 33] = data.Expertise.CountCaseEkmpTarget;
                ObjWorkSheet.Cells[currentIndex, 34] = data.Expertise.CountCaseEkmpPlan;
                ObjWorkSheet.Cells[currentIndex, 36] = data.Expertise.CountCaseEkmpComplaint;
                ObjWorkSheet.Cells[currentIndex, 37] = data.Expertise.CountCaseEkmpLeth;
                ObjWorkSheet.Cells[currentIndex, 38] = data.Expertise.CountCaseEkmpByMek;
                ObjWorkSheet.Cells[currentIndex, 39] = data.Expertise.CountCaseEkmpByMee;
                ObjWorkSheet.Cells[currentIndex, 40] = data.Expertise.CountCaseEkmpUTheme;
                ObjWorkSheet.Cells[currentIndex, 41] = data.Expertise.CountCaseEkmpMultiTarget;
                ObjWorkSheet.Cells[currentIndex, 42] = data.Expertise.CountCaseEkmpMultiPlan;
                ObjWorkSheet.Cells[currentIndex, 44] = data.Expertise.CountCaseEkmpMultiLeth;
                ObjWorkSheet.Cells[currentIndex, 45] = data.Expertise.CountCaseEkmpMultiUthemeTarget;
                ObjWorkSheet.Cells[currentIndex, 46] = data.Expertise.CountCaseEkmpMultiUthemePlan;
                ObjWorkSheet.Cells[currentIndex, 48] = data.Expertise.CountCaseDefectedBySmoTarget;
                ObjWorkSheet.Cells[currentIndex, 49] = data.Expertise.CountCaseDefectedBySmoPlan;
                ObjWorkSheet.Cells[currentIndex, 51] = data.Expertise.CountEkmpDefectedCaseTarget;
                ObjWorkSheet.Cells[currentIndex, 52] = data.Expertise.CountEkmpDefectedCasePlan;
                ObjWorkSheet.Cells[currentIndex, 54] = data.Expertise.CountEkmpBadTarget;
                ObjWorkSheet.Cells[currentIndex, 55] = data.Expertise.CountEkmpBadPlan;
                ObjWorkSheet.Cells[currentIndex, 57] = data.Expertise.CountEkmpBadDs;
                ObjWorkSheet.Cells[currentIndex, 58] = data.Expertise.CountEkmpBadDsNotAffected;
                ObjWorkSheet.Cells[currentIndex, 59] = data.Expertise.CountEkmpBadDsProlonger;
                ObjWorkSheet.Cells[currentIndex, 60] = data.Expertise.CountEkmpBadDsDecline;
                ObjWorkSheet.Cells[currentIndex, 61] = data.Expertise.CountEkmpBadDsInjured;
                ObjWorkSheet.Cells[currentIndex, 62] = data.Expertise.CountEkmpBadDsLeth;
                ObjWorkSheet.Cells[currentIndex, 63] = data.Expertise.CountEkmpBadMed;
                ObjWorkSheet.Cells[currentIndex, 64] = data.Expertise.CountEkmpUnreglamentedMed;
                ObjWorkSheet.Cells[currentIndex, 65] = data.Expertise.CountEkmpStopMed;
                ObjWorkSheet.Cells[currentIndex, 66] = data.Expertise.CountEkmpContinuity;
                ObjWorkSheet.Cells[currentIndex, 67] = data.Expertise.CountEkmpUnprofile;
                ObjWorkSheet.Cells[currentIndex, 68] = data.Expertise.CountEkmpUnfounded;
                ObjWorkSheet.Cells[currentIndex, 69] = data.Expertise.CountEkmpRepeat;
                ObjWorkSheet.Cells[currentIndex, 70] = data.Expertise.CountEkmpDifference;
                ObjWorkSheet.Cells[currentIndex, 71] = data.Expertise.CountEkmpUnfoundedMedicaments;
                ObjWorkSheet.Cells[currentIndex, 72] = data.Expertise.CountEkmpUnfoundedReject;
                ObjWorkSheet.Cells[currentIndex, 73] = data.Expertise.CountEkmpDisp;
                ObjWorkSheet.Cells[currentIndex, 74] = data.Expertise.CountEkmpRepeat2weeks;
                ObjWorkSheet.Cells[currentIndex, 75] = data.Expertise.CountEkmpOutOfResults;
                ObjWorkSheet.Cells[currentIndex++, 76] = data.Expertise.CountEkmpDoubleHospital;

            }
        }
    }
}
