using System;
using System.Collections.Generic;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelControlZpz2023FullCreator : ExcelBaseCreator<CReportZpz2023Full[]>
    {
        private const int StartPositionExpertises = 6;
        private const int StartPositionFinance = 7;
        private string _year;

        public ExcelControlZpz2023FullCreator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.ControlZpz2023Full, header, filialName, false) { }

        protected override void FillReport(CReportZpz2023Full[] reports, CReportZpz2023Full[] yearReport)
        {
            FillExpertises(reports);
            FillFinance(reports);
        }

        

        private void FillFinance(CReportZpz2023Full[] reports)
        {
            var finance = reports.Select(r => new { r.Filial, r.Finance1Q, r.Finance2Q, r.Finance3Q, r.Finance4Q }).ToList();

            int countReport = finance.Count;
            int currentIndex = StartPositionFinance;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[2];
            //CopyNullCells(ObjWorkSheet, countReport, StartPositionFinance);

            int counter = 1;
            foreach (var data in finance)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Finance1Q.SumPayment;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Finance1Q.SumNotPayment;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Finance1Q.SumMek;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Finance1Q.SumMee;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Finance1Q.SumEkmp;

                ObjWorkSheet.Cells[currentIndex, 9] = data.Finance2Q.SumPayment;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Finance2Q.SumNotPayment;
                ObjWorkSheet.Cells[currentIndex, 11] = data.Finance2Q.SumMek;
                ObjWorkSheet.Cells[currentIndex, 12] = data.Finance2Q.SumMee;
                ObjWorkSheet.Cells[currentIndex, 13] = data.Finance2Q.SumEkmp;

                ObjWorkSheet.Cells[currentIndex, 15] = data.Finance3Q.SumPayment;
                ObjWorkSheet.Cells[currentIndex, 16] = data.Finance3Q.SumNotPayment;
                ObjWorkSheet.Cells[currentIndex, 17] = data.Finance3Q.SumMek;
                ObjWorkSheet.Cells[currentIndex, 18] = data.Finance3Q.SumMee;
                ObjWorkSheet.Cells[currentIndex, 19] = data.Finance3Q.SumEkmp;

                ObjWorkSheet.Cells[currentIndex, 21] = data.Finance4Q.SumPayment;
                ObjWorkSheet.Cells[currentIndex, 22] = data.Finance4Q.SumNotPayment;
                ObjWorkSheet.Cells[currentIndex, 23] = data.Finance4Q.SumMek;
                ObjWorkSheet.Cells[currentIndex, 24] = data.Finance4Q.SumMee;
                ObjWorkSheet.Cells[currentIndex++, 25] = data.Finance4Q.SumEkmp;
            }
        }

        private void FillExpertises(CReportZpz2023Full[] reports)
        {
            var expertises = reports.Select(r => new { r.Filial, r.Expertise1Q, r.Expertise2Q, r.Expertise3Q, r.Expertise4Q }).ToList();

            int countReport = expertises.Count;
            int currentIndex = StartPositionExpertises;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            //CopyNullCells4Rows(ObjWorkSheet, countReport, StartPositionExpertises);

            foreach (var data in expertises) 
            {
                    ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                    ObjWorkSheet.Cells[currentIndex, 3] = data.Expertise1Q.Bills;
                    ObjWorkSheet.Cells[currentIndex, 4] = data.Expertise1Q.CountMeeTarget;
                    ObjWorkSheet.Cells[currentIndex, 5] = data.Expertise1Q.CountMeePlan;
                    ObjWorkSheet.Cells[currentIndex, 7] = data.Expertise1Q.CountMeeComplaintTarget;
                    ObjWorkSheet.Cells[currentIndex, 8] = data.Expertise1Q.CountMeeComplaintPlan;
                    ObjWorkSheet.Cells[currentIndex, 10] = data.Expertise1Q.CountMeeRepeat;
                    ObjWorkSheet.Cells[currentIndex, 11] = data.Expertise1Q.CountMeeOnco;
                    ObjWorkSheet.Cells[currentIndex, 12] = data.Expertise1Q.CountMeeDs;
                    ObjWorkSheet.Cells[currentIndex, 13] = data.Expertise1Q.CountMeeLeth;
                    ObjWorkSheet.Cells[currentIndex, 14] = data.Expertise1Q.CountMeeInjured;
                    ObjWorkSheet.Cells[currentIndex, 15] = data.Expertise1Q.CountMeeDefectedCaseTarget;
                    ObjWorkSheet.Cells[currentIndex, 16] = data.Expertise1Q.CountMeeDefectedCasePlan;
                    ObjWorkSheet.Cells[currentIndex, 18] = data.Expertise1Q.CountMeeDefectsTarget;
                    ObjWorkSheet.Cells[currentIndex, 19] = data.Expertise1Q.CountMeeDefectsPlan;
                    ObjWorkSheet.Cells[currentIndex, 21] = data.Expertise1Q.CountMeeDefectsPeriod;
                    ObjWorkSheet.Cells[currentIndex, 22] = data.Expertise1Q.CountMeeDefectsCondition;
                    ObjWorkSheet.Cells[currentIndex, 23] = data.Expertise1Q.CountMeeDefectsRepeat;
                    ObjWorkSheet.Cells[currentIndex, 24] = data.Expertise1Q.CountMeeDefectsOutOfDocums;
                    ObjWorkSheet.Cells[currentIndex, 25] = data.Expertise1Q.CountMeeDefectsUnpayable;
                    ObjWorkSheet.Cells[currentIndex, 26] = data.Expertise1Q.CountMeeDefectsBuyMedicament;
                    ObjWorkSheet.Cells[currentIndex, 27] = data.Expertise1Q.CountMeeDefectsOutOfLeth;
                    ObjWorkSheet.Cells[currentIndex, 28] = data.Expertise1Q.CountMeeDefectsWithoutDocums;
                    ObjWorkSheet.Cells[currentIndex, 29] = data.Expertise1Q.CountMeeDefectsIncorrectDocums;
                    ObjWorkSheet.Cells[currentIndex, 30] = data.Expertise1Q.CountMeeDefectsBadDocums;
                    ObjWorkSheet.Cells[currentIndex, 31] = data.Expertise1Q.CountMeeDefectsBadDate;
                    ObjWorkSheet.Cells[currentIndex, 32] = data.Expertise1Q.CountMeeDefectsBadData;
                    ObjWorkSheet.Cells[currentIndex, 33] = data.Expertise1Q.CountMeeDefectsOutOfProtocol;
                    ObjWorkSheet.Cells[currentIndex, 34] = data.Expertise1Q.CountCaseEkmpTarget;
                    ObjWorkSheet.Cells[currentIndex, 35] = data.Expertise1Q.CountCaseEkmpPlan;
                    ObjWorkSheet.Cells[currentIndex, 37] = data.Expertise1Q.CountCaseEkmpComplaint;
                    ObjWorkSheet.Cells[currentIndex, 38] = data.Expertise1Q.CountCaseEkmpLeth;
                    ObjWorkSheet.Cells[currentIndex, 39] = data.Expertise1Q.CountCaseEkmpByMek;
                    ObjWorkSheet.Cells[currentIndex, 40] = data.Expertise1Q.CountCaseEkmpByMee;
                    ObjWorkSheet.Cells[currentIndex, 41] = data.Expertise1Q.CountCaseEkmpUTheme;
                    ObjWorkSheet.Cells[currentIndex, 42] = data.Expertise1Q.CountCaseEkmpMultiTarget;
                    ObjWorkSheet.Cells[currentIndex, 43] = data.Expertise1Q.CountCaseEkmpMultiPlan;
                    ObjWorkSheet.Cells[currentIndex, 45] = data.Expertise1Q.CountCaseEkmpMultiLeth;
                    ObjWorkSheet.Cells[currentIndex, 46] = data.Expertise1Q.CountCaseEkmpMultiUthemeTarget;
                    ObjWorkSheet.Cells[currentIndex, 47] = data.Expertise1Q.CountCaseEkmpMultiUthemePlan;
                    ObjWorkSheet.Cells[currentIndex, 49] = data.Expertise1Q.CountCaseDefectedBySmoTarget;
                    ObjWorkSheet.Cells[currentIndex, 50] = data.Expertise1Q.CountCaseDefectedBySmoPlan;
                    ObjWorkSheet.Cells[currentIndex, 52] = data.Expertise1Q.CountEkmpDefectedCaseTarget;
                    ObjWorkSheet.Cells[currentIndex, 53] = data.Expertise1Q.CountEkmpDefectedCasePlan;
                    ObjWorkSheet.Cells[currentIndex, 55] = data.Expertise1Q.CountEkmpBadDs;
                    ObjWorkSheet.Cells[currentIndex, 56] = data.Expertise1Q.CountEkmpBadDsNotAffected;
                    ObjWorkSheet.Cells[currentIndex, 57] = data.Expertise1Q.CountEkmpBadDsProlonger;
                    ObjWorkSheet.Cells[currentIndex, 58] = data.Expertise1Q.CountEkmpBadDsDecline;
                    ObjWorkSheet.Cells[currentIndex, 59] = data.Expertise1Q.CountEkmpBadDsInjured;
                    ObjWorkSheet.Cells[currentIndex, 60] = data.Expertise1Q.CountEkmpBadDsLeth;
                    ObjWorkSheet.Cells[currentIndex, 61] = data.Expertise1Q.CountEkmpBadMed;
                    ObjWorkSheet.Cells[currentIndex, 62] = data.Expertise1Q.CountEkmpUnreglamentedMed;
                    ObjWorkSheet.Cells[currentIndex, 63] = data.Expertise1Q.CountEkmpStopMed;
                    ObjWorkSheet.Cells[currentIndex, 64] = data.Expertise1Q.CountEkmpContinuity;
                    ObjWorkSheet.Cells[currentIndex, 65] = data.Expertise1Q.CountEkmpUnprofile;
                    ObjWorkSheet.Cells[currentIndex, 66] = data.Expertise1Q.CountEkmpUnfounded;
                    ObjWorkSheet.Cells[currentIndex, 67] = data.Expertise1Q.CountEkmpRepeat;
                    ObjWorkSheet.Cells[currentIndex, 68] = data.Expertise1Q.CountEkmpDifference;
                    ObjWorkSheet.Cells[currentIndex, 69] = data.Expertise1Q.CountEkmpUnfoundedMedicaments;
                    ObjWorkSheet.Cells[currentIndex, 70] = data.Expertise1Q.CountEkmpUnfoundedReject;
                    ObjWorkSheet.Cells[currentIndex, 71] = data.Expertise1Q.CountEkmpDisp;
                    ObjWorkSheet.Cells[currentIndex, 72] = data.Expertise1Q.CountEkmpRepeat2weeks;
                    ObjWorkSheet.Cells[currentIndex, 73] = data.Expertise1Q.CountEkmpOutOfResults;
                    ObjWorkSheet.Cells[currentIndex++, 74] = data.Expertise1Q.CountEkmpDoubleHospital;
                    ObjWorkSheet.Cells[currentIndex, 3] = data.Expertise2Q.Bills;
                    ObjWorkSheet.Cells[currentIndex, 4] = data.Expertise2Q.CountMeeTarget;
                    ObjWorkSheet.Cells[currentIndex, 5] = data.Expertise2Q.CountMeePlan;
                    ObjWorkSheet.Cells[currentIndex, 7] = data.Expertise2Q.CountMeeComplaintTarget;
                    ObjWorkSheet.Cells[currentIndex, 8] = data.Expertise2Q.CountMeeComplaintPlan;
                    ObjWorkSheet.Cells[currentIndex, 10] = data.Expertise2Q.CountMeeRepeat;
                    ObjWorkSheet.Cells[currentIndex, 11] = data.Expertise2Q.CountMeeOnco;
                    ObjWorkSheet.Cells[currentIndex, 12] = data.Expertise2Q.CountMeeDs;
                    ObjWorkSheet.Cells[currentIndex, 13] = data.Expertise2Q.CountMeeLeth;
                    ObjWorkSheet.Cells[currentIndex, 14] = data.Expertise2Q.CountMeeInjured;
                    ObjWorkSheet.Cells[currentIndex, 15] = data.Expertise2Q.CountMeeDefectedCaseTarget;
                    ObjWorkSheet.Cells[currentIndex, 16] = data.Expertise2Q.CountMeeDefectedCasePlan;
                    ObjWorkSheet.Cells[currentIndex, 18] = data.Expertise2Q.CountMeeDefectsTarget;
                    ObjWorkSheet.Cells[currentIndex, 19] = data.Expertise2Q.CountMeeDefectsPlan;
                    ObjWorkSheet.Cells[currentIndex, 21] = data.Expertise2Q.CountMeeDefectsPeriod;
                    ObjWorkSheet.Cells[currentIndex, 22] = data.Expertise2Q.CountMeeDefectsCondition;
                    ObjWorkSheet.Cells[currentIndex, 23] = data.Expertise2Q.CountMeeDefectsRepeat;
                    ObjWorkSheet.Cells[currentIndex, 24] = data.Expertise2Q.CountMeeDefectsOutOfDocums;
                    ObjWorkSheet.Cells[currentIndex, 25] = data.Expertise2Q.CountMeeDefectsUnpayable;
                    ObjWorkSheet.Cells[currentIndex, 26] = data.Expertise2Q.CountMeeDefectsBuyMedicament;
                    ObjWorkSheet.Cells[currentIndex, 27] = data.Expertise2Q.CountMeeDefectsOutOfLeth;
                    ObjWorkSheet.Cells[currentIndex, 28] = data.Expertise2Q.CountMeeDefectsWithoutDocums;
                    ObjWorkSheet.Cells[currentIndex, 29] = data.Expertise2Q.CountMeeDefectsIncorrectDocums;
                    ObjWorkSheet.Cells[currentIndex, 30] = data.Expertise2Q.CountMeeDefectsBadDocums;
                    ObjWorkSheet.Cells[currentIndex, 31] = data.Expertise2Q.CountMeeDefectsBadDate;
                    ObjWorkSheet.Cells[currentIndex, 32] = data.Expertise2Q.CountMeeDefectsBadData;
                    ObjWorkSheet.Cells[currentIndex, 33] = data.Expertise2Q.CountMeeDefectsOutOfProtocol;
                    ObjWorkSheet.Cells[currentIndex, 34] = data.Expertise2Q.CountCaseEkmpTarget;
                    ObjWorkSheet.Cells[currentIndex, 35] = data.Expertise2Q.CountCaseEkmpPlan;
                    ObjWorkSheet.Cells[currentIndex, 37] = data.Expertise2Q.CountCaseEkmpComplaint;
                    ObjWorkSheet.Cells[currentIndex, 38] = data.Expertise2Q.CountCaseEkmpLeth;
                    ObjWorkSheet.Cells[currentIndex, 39] = data.Expertise2Q.CountCaseEkmpByMek;
                    ObjWorkSheet.Cells[currentIndex, 40] = data.Expertise2Q.CountCaseEkmpByMee;
                    ObjWorkSheet.Cells[currentIndex, 41] = data.Expertise2Q.CountCaseEkmpUTheme;
                    ObjWorkSheet.Cells[currentIndex, 42] = data.Expertise2Q.CountCaseEkmpMultiTarget;
                    ObjWorkSheet.Cells[currentIndex, 43] = data.Expertise2Q.CountCaseEkmpMultiPlan;
                    ObjWorkSheet.Cells[currentIndex, 45] = data.Expertise2Q.CountCaseEkmpMultiLeth;
                    ObjWorkSheet.Cells[currentIndex, 46] = data.Expertise2Q.CountCaseEkmpMultiUthemeTarget;
                    ObjWorkSheet.Cells[currentIndex, 47] = data.Expertise2Q.CountCaseEkmpMultiUthemePlan;
                    ObjWorkSheet.Cells[currentIndex, 49] = data.Expertise2Q.CountCaseDefectedBySmoTarget;
                    ObjWorkSheet.Cells[currentIndex, 50] = data.Expertise2Q.CountCaseDefectedBySmoPlan;
                    ObjWorkSheet.Cells[currentIndex, 52] = data.Expertise2Q.CountEkmpDefectedCaseTarget;
                    ObjWorkSheet.Cells[currentIndex, 53] = data.Expertise2Q.CountEkmpDefectedCasePlan;
                    ObjWorkSheet.Cells[currentIndex, 55] = data.Expertise2Q.CountEkmpBadDs;
                    ObjWorkSheet.Cells[currentIndex, 56] = data.Expertise2Q.CountEkmpBadDsNotAffected;
                    ObjWorkSheet.Cells[currentIndex, 57] = data.Expertise2Q.CountEkmpBadDsProlonger;
                    ObjWorkSheet.Cells[currentIndex, 58] = data.Expertise2Q.CountEkmpBadDsDecline;
                    ObjWorkSheet.Cells[currentIndex, 59] = data.Expertise2Q.CountEkmpBadDsInjured;
                    ObjWorkSheet.Cells[currentIndex, 60] = data.Expertise2Q.CountEkmpBadDsLeth;
                    ObjWorkSheet.Cells[currentIndex, 61] = data.Expertise2Q.CountEkmpBadMed;
                    ObjWorkSheet.Cells[currentIndex, 62] = data.Expertise2Q.CountEkmpUnreglamentedMed;
                    ObjWorkSheet.Cells[currentIndex, 63] = data.Expertise2Q.CountEkmpStopMed;
                    ObjWorkSheet.Cells[currentIndex, 64] = data.Expertise2Q.CountEkmpContinuity;
                    ObjWorkSheet.Cells[currentIndex, 65] = data.Expertise2Q.CountEkmpUnprofile;
                    ObjWorkSheet.Cells[currentIndex, 66] = data.Expertise2Q.CountEkmpUnfounded;
                    ObjWorkSheet.Cells[currentIndex, 67] = data.Expertise2Q.CountEkmpRepeat;
                    ObjWorkSheet.Cells[currentIndex, 68] = data.Expertise2Q.CountEkmpDifference;
                    ObjWorkSheet.Cells[currentIndex, 69] = data.Expertise2Q.CountEkmpUnfoundedMedicaments;
                    ObjWorkSheet.Cells[currentIndex, 70] = data.Expertise2Q.CountEkmpUnfoundedReject;
                    ObjWorkSheet.Cells[currentIndex, 71] = data.Expertise2Q.CountEkmpDisp;
                    ObjWorkSheet.Cells[currentIndex, 72] = data.Expertise2Q.CountEkmpRepeat2weeks;
                    ObjWorkSheet.Cells[currentIndex, 73] = data.Expertise2Q.CountEkmpOutOfResults;
                    ObjWorkSheet.Cells[currentIndex++, 74] = data.Expertise2Q.CountEkmpDoubleHospital;
                    ObjWorkSheet.Cells[currentIndex, 3] = data.Expertise3Q.Bills;
                    ObjWorkSheet.Cells[currentIndex, 4] = data.Expertise3Q.CountMeeTarget;
                    ObjWorkSheet.Cells[currentIndex, 5] = data.Expertise3Q.CountMeePlan;
                    ObjWorkSheet.Cells[currentIndex, 7] = data.Expertise3Q.CountMeeComplaintTarget;
                    ObjWorkSheet.Cells[currentIndex, 8] = data.Expertise3Q.CountMeeComplaintPlan;
                    ObjWorkSheet.Cells[currentIndex, 10] = data.Expertise3Q.CountMeeRepeat;
                    ObjWorkSheet.Cells[currentIndex, 11] = data.Expertise3Q.CountMeeOnco;
                    ObjWorkSheet.Cells[currentIndex, 12] = data.Expertise3Q.CountMeeDs;
                    ObjWorkSheet.Cells[currentIndex, 13] = data.Expertise3Q.CountMeeLeth;
                    ObjWorkSheet.Cells[currentIndex, 14] = data.Expertise3Q.CountMeeInjured;
                    ObjWorkSheet.Cells[currentIndex, 15] = data.Expertise3Q.CountMeeDefectedCaseTarget;
                    ObjWorkSheet.Cells[currentIndex, 16] = data.Expertise3Q.CountMeeDefectedCasePlan;
                    ObjWorkSheet.Cells[currentIndex, 18] = data.Expertise3Q.CountMeeDefectsTarget;
                    ObjWorkSheet.Cells[currentIndex, 19] = data.Expertise3Q.CountMeeDefectsPlan;
                    ObjWorkSheet.Cells[currentIndex, 21] = data.Expertise3Q.CountMeeDefectsPeriod;
                    ObjWorkSheet.Cells[currentIndex, 22] = data.Expertise3Q.CountMeeDefectsCondition;
                    ObjWorkSheet.Cells[currentIndex, 23] = data.Expertise3Q.CountMeeDefectsRepeat;
                    ObjWorkSheet.Cells[currentIndex, 24] = data.Expertise3Q.CountMeeDefectsOutOfDocums;
                    ObjWorkSheet.Cells[currentIndex, 25] = data.Expertise3Q.CountMeeDefectsUnpayable;
                    ObjWorkSheet.Cells[currentIndex, 26] = data.Expertise3Q.CountMeeDefectsBuyMedicament;
                    ObjWorkSheet.Cells[currentIndex, 27] = data.Expertise3Q.CountMeeDefectsOutOfLeth;
                    ObjWorkSheet.Cells[currentIndex, 28] = data.Expertise3Q.CountMeeDefectsWithoutDocums;
                    ObjWorkSheet.Cells[currentIndex, 29] = data.Expertise3Q.CountMeeDefectsIncorrectDocums;
                    ObjWorkSheet.Cells[currentIndex, 30] = data.Expertise3Q.CountMeeDefectsBadDocums;
                    ObjWorkSheet.Cells[currentIndex, 31] = data.Expertise3Q.CountMeeDefectsBadDate;
                    ObjWorkSheet.Cells[currentIndex, 32] = data.Expertise3Q.CountMeeDefectsBadData;
                    ObjWorkSheet.Cells[currentIndex, 33] = data.Expertise3Q.CountMeeDefectsOutOfProtocol;
                    ObjWorkSheet.Cells[currentIndex, 34] = data.Expertise3Q.CountCaseEkmpTarget;
                    ObjWorkSheet.Cells[currentIndex, 35] = data.Expertise3Q.CountCaseEkmpPlan;
                    ObjWorkSheet.Cells[currentIndex, 37] = data.Expertise3Q.CountCaseEkmpComplaint;
                    ObjWorkSheet.Cells[currentIndex, 38] = data.Expertise3Q.CountCaseEkmpLeth;
                    ObjWorkSheet.Cells[currentIndex, 39] = data.Expertise3Q.CountCaseEkmpByMek;
                    ObjWorkSheet.Cells[currentIndex, 40] = data.Expertise3Q.CountCaseEkmpByMee;
                    ObjWorkSheet.Cells[currentIndex, 41] = data.Expertise3Q.CountCaseEkmpUTheme;
                    ObjWorkSheet.Cells[currentIndex, 42] = data.Expertise3Q.CountCaseEkmpMultiTarget;
                    ObjWorkSheet.Cells[currentIndex, 43] = data.Expertise3Q.CountCaseEkmpMultiPlan;
                    ObjWorkSheet.Cells[currentIndex, 45] = data.Expertise3Q.CountCaseEkmpMultiLeth;
                    ObjWorkSheet.Cells[currentIndex, 46] = data.Expertise3Q.CountCaseEkmpMultiUthemeTarget;
                    ObjWorkSheet.Cells[currentIndex, 47] = data.Expertise3Q.CountCaseEkmpMultiUthemePlan;
                    ObjWorkSheet.Cells[currentIndex, 49] = data.Expertise3Q.CountCaseDefectedBySmoTarget;
                    ObjWorkSheet.Cells[currentIndex, 50] = data.Expertise3Q.CountCaseDefectedBySmoPlan;
                    ObjWorkSheet.Cells[currentIndex, 52] = data.Expertise3Q.CountEkmpDefectedCaseTarget;
                    ObjWorkSheet.Cells[currentIndex, 53] = data.Expertise3Q.CountEkmpDefectedCasePlan;
                    ObjWorkSheet.Cells[currentIndex, 55] = data.Expertise3Q.CountEkmpBadDs;
                    ObjWorkSheet.Cells[currentIndex, 56] = data.Expertise3Q.CountEkmpBadDsNotAffected;
                    ObjWorkSheet.Cells[currentIndex, 57] = data.Expertise3Q.CountEkmpBadDsProlonger;
                    ObjWorkSheet.Cells[currentIndex, 58] = data.Expertise3Q.CountEkmpBadDsDecline;
                    ObjWorkSheet.Cells[currentIndex, 59] = data.Expertise3Q.CountEkmpBadDsInjured;
                    ObjWorkSheet.Cells[currentIndex, 60] = data.Expertise3Q.CountEkmpBadDsLeth;
                    ObjWorkSheet.Cells[currentIndex, 61] = data.Expertise3Q.CountEkmpBadMed;
                    ObjWorkSheet.Cells[currentIndex, 62] = data.Expertise3Q.CountEkmpUnreglamentedMed;
                    ObjWorkSheet.Cells[currentIndex, 63] = data.Expertise3Q.CountEkmpStopMed;
                    ObjWorkSheet.Cells[currentIndex, 64] = data.Expertise3Q.CountEkmpContinuity;
                    ObjWorkSheet.Cells[currentIndex, 65] = data.Expertise3Q.CountEkmpUnprofile;
                    ObjWorkSheet.Cells[currentIndex, 66] = data.Expertise3Q.CountEkmpUnfounded;
                    ObjWorkSheet.Cells[currentIndex, 67] = data.Expertise3Q.CountEkmpRepeat;
                    ObjWorkSheet.Cells[currentIndex, 68] = data.Expertise3Q.CountEkmpDifference;
                    ObjWorkSheet.Cells[currentIndex, 69] = data.Expertise3Q.CountEkmpUnfoundedMedicaments;
                    ObjWorkSheet.Cells[currentIndex, 70] = data.Expertise3Q.CountEkmpUnfoundedReject;
                    ObjWorkSheet.Cells[currentIndex, 71] = data.Expertise3Q.CountEkmpDisp;
                    ObjWorkSheet.Cells[currentIndex, 72] = data.Expertise3Q.CountEkmpRepeat2weeks;
                    ObjWorkSheet.Cells[currentIndex, 73] = data.Expertise3Q.CountEkmpOutOfResults;
                    ObjWorkSheet.Cells[currentIndex++, 74] = data.Expertise3Q.CountEkmpDoubleHospital;
                    ObjWorkSheet.Cells[currentIndex, 3] = data.Expertise4Q.Bills;
                    ObjWorkSheet.Cells[currentIndex, 4] = data.Expertise4Q.CountMeeTarget;
                    ObjWorkSheet.Cells[currentIndex, 5] = data.Expertise4Q.CountMeePlan;
                    ObjWorkSheet.Cells[currentIndex, 7] = data.Expertise4Q.CountMeeComplaintTarget;
                    ObjWorkSheet.Cells[currentIndex, 8] = data.Expertise4Q.CountMeeComplaintPlan;
                    ObjWorkSheet.Cells[currentIndex, 10] = data.Expertise4Q.CountMeeRepeat;
                    ObjWorkSheet.Cells[currentIndex, 11] = data.Expertise4Q.CountMeeOnco;
                    ObjWorkSheet.Cells[currentIndex, 12] = data.Expertise4Q.CountMeeDs;
                    ObjWorkSheet.Cells[currentIndex, 13] = data.Expertise4Q.CountMeeLeth;
                    ObjWorkSheet.Cells[currentIndex, 14] = data.Expertise4Q.CountMeeInjured;
                    ObjWorkSheet.Cells[currentIndex, 15] = data.Expertise4Q.CountMeeDefectedCaseTarget;
                    ObjWorkSheet.Cells[currentIndex, 16] = data.Expertise4Q.CountMeeDefectedCasePlan;
                    ObjWorkSheet.Cells[currentIndex, 18] = data.Expertise4Q.CountMeeDefectsTarget;
                    ObjWorkSheet.Cells[currentIndex, 19] = data.Expertise4Q.CountMeeDefectsPlan;
                    ObjWorkSheet.Cells[currentIndex, 21] = data.Expertise4Q.CountMeeDefectsPeriod;
                    ObjWorkSheet.Cells[currentIndex, 22] = data.Expertise4Q.CountMeeDefectsCondition;
                    ObjWorkSheet.Cells[currentIndex, 23] = data.Expertise4Q.CountMeeDefectsRepeat;
                    ObjWorkSheet.Cells[currentIndex, 24] = data.Expertise4Q.CountMeeDefectsOutOfDocums;
                    ObjWorkSheet.Cells[currentIndex, 25] = data.Expertise4Q.CountMeeDefectsUnpayable;
                    ObjWorkSheet.Cells[currentIndex, 26] = data.Expertise4Q.CountMeeDefectsBuyMedicament;
                    ObjWorkSheet.Cells[currentIndex, 27] = data.Expertise4Q.CountMeeDefectsOutOfLeth;
                    ObjWorkSheet.Cells[currentIndex, 28] = data.Expertise4Q.CountMeeDefectsWithoutDocums;
                    ObjWorkSheet.Cells[currentIndex, 29] = data.Expertise4Q.CountMeeDefectsIncorrectDocums;
                    ObjWorkSheet.Cells[currentIndex, 30] = data.Expertise4Q.CountMeeDefectsBadDocums;
                    ObjWorkSheet.Cells[currentIndex, 31] = data.Expertise4Q.CountMeeDefectsBadDate;
                    ObjWorkSheet.Cells[currentIndex, 32] = data.Expertise4Q.CountMeeDefectsBadData;
                    ObjWorkSheet.Cells[currentIndex, 33] = data.Expertise4Q.CountMeeDefectsOutOfProtocol;
                    ObjWorkSheet.Cells[currentIndex, 34] = data.Expertise4Q.CountCaseEkmpTarget;
                    ObjWorkSheet.Cells[currentIndex, 35] = data.Expertise4Q.CountCaseEkmpPlan;
                    ObjWorkSheet.Cells[currentIndex, 37] = data.Expertise4Q.CountCaseEkmpComplaint;
                    ObjWorkSheet.Cells[currentIndex, 38] = data.Expertise4Q.CountCaseEkmpLeth;
                    ObjWorkSheet.Cells[currentIndex, 39] = data.Expertise4Q.CountCaseEkmpByMek;
                    ObjWorkSheet.Cells[currentIndex, 40] = data.Expertise4Q.CountCaseEkmpByMee;
                    ObjWorkSheet.Cells[currentIndex, 41] = data.Expertise4Q.CountCaseEkmpUTheme;
                    ObjWorkSheet.Cells[currentIndex, 42] = data.Expertise4Q.CountCaseEkmpMultiTarget;
                    ObjWorkSheet.Cells[currentIndex, 43] = data.Expertise4Q.CountCaseEkmpMultiPlan;
                    ObjWorkSheet.Cells[currentIndex, 45] = data.Expertise4Q.CountCaseEkmpMultiLeth;
                    ObjWorkSheet.Cells[currentIndex, 46] = data.Expertise4Q.CountCaseEkmpMultiUthemeTarget;
                    ObjWorkSheet.Cells[currentIndex, 47] = data.Expertise4Q.CountCaseEkmpMultiUthemePlan;
                    ObjWorkSheet.Cells[currentIndex, 49] = data.Expertise4Q.CountCaseDefectedBySmoTarget;
                    ObjWorkSheet.Cells[currentIndex, 50] = data.Expertise4Q.CountCaseDefectedBySmoPlan;
                    ObjWorkSheet.Cells[currentIndex, 52] = data.Expertise4Q.CountEkmpDefectedCaseTarget;
                    ObjWorkSheet.Cells[currentIndex, 53] = data.Expertise4Q.CountEkmpDefectedCasePlan;
                    ObjWorkSheet.Cells[currentIndex, 55] = data.Expertise4Q.CountEkmpBadDs;
                    ObjWorkSheet.Cells[currentIndex, 56] = data.Expertise4Q.CountEkmpBadDsNotAffected;
                    ObjWorkSheet.Cells[currentIndex, 57] = data.Expertise4Q.CountEkmpBadDsProlonger;
                    ObjWorkSheet.Cells[currentIndex, 58] = data.Expertise4Q.CountEkmpBadDsDecline;
                    ObjWorkSheet.Cells[currentIndex, 59] = data.Expertise4Q.CountEkmpBadDsInjured;
                    ObjWorkSheet.Cells[currentIndex, 60] = data.Expertise4Q.CountEkmpBadDsLeth;
                    ObjWorkSheet.Cells[currentIndex, 61] = data.Expertise4Q.CountEkmpBadMed;
                    ObjWorkSheet.Cells[currentIndex, 62] = data.Expertise4Q.CountEkmpUnreglamentedMed;
                    ObjWorkSheet.Cells[currentIndex, 63] = data.Expertise4Q.CountEkmpStopMed;
                    ObjWorkSheet.Cells[currentIndex, 64] = data.Expertise4Q.CountEkmpContinuity;
                    ObjWorkSheet.Cells[currentIndex, 65] = data.Expertise4Q.CountEkmpUnprofile;
                    ObjWorkSheet.Cells[currentIndex, 66] = data.Expertise4Q.CountEkmpUnfounded;
                    ObjWorkSheet.Cells[currentIndex, 67] = data.Expertise4Q.CountEkmpRepeat;
                    ObjWorkSheet.Cells[currentIndex, 68] = data.Expertise4Q.CountEkmpDifference;
                    ObjWorkSheet.Cells[currentIndex, 69] = data.Expertise4Q.CountEkmpUnfoundedMedicaments;
                    ObjWorkSheet.Cells[currentIndex, 70] = data.Expertise4Q.CountEkmpUnfoundedReject;
                    ObjWorkSheet.Cells[currentIndex, 71] = data.Expertise4Q.CountEkmpDisp;
                    ObjWorkSheet.Cells[currentIndex, 72] = data.Expertise4Q.CountEkmpRepeat2weeks;
                    ObjWorkSheet.Cells[currentIndex, 73] = data.Expertise4Q.CountEkmpOutOfResults;
                    ObjWorkSheet.Cells[currentIndex++, 74] = data.Expertise4Q.CountEkmpDoubleHospital;
            }
        }
    }
}
