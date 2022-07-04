using System;
using System.Collections.Generic;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelControlZpzCreator : ExcelBaseCreator<CReportPg[]>
    {
        private const int StartPosition = 6;
        private const int StartPersonnel = 7;
        private string _yymm;

        public ExcelControlZpzCreator(
            string filename,
            string header,
            string filialName, string yymm) : base(filename, ExcelForm.ControlZpz, header, filialName, false) { _yymm = yymm; }

        protected override void FillReport(CReportPg[] reports, CReportPg[] yearReport)
        {
            FillExpertises(reports);
            FillFinance(reports);
            FillPersonnel(reports);
            FillNormative(reports);
        }

        private void FillNormative(CReportPg[] reports)
        {
            //Dictionary<int, double> oldNormative = new Dictionary<int, double>()
            //{
            //    { 6, 3},
            //    { 10, 0.8},
            //    { 14, 8},
            //    { 18, 8},
            //    { 21, 1.5},
            //    { 24, 0.5},
            //    { 24, 0.5},
            //    { 24, 3},
            //    { 30, 5},
            //};


            Dictionary<int, double> newNormative = new Dictionary<int, double>()
            {
                { 6, 2},
                { 10, 0.5},
                { 14, 6},
                { 18, 6},
                { 21, 0.5},
                { 24, 0.2},
                { 27, 1.5},
                { 30, 3},
            };

            var normative = reports.Select(r => new { r.Filial, r.Normative }).ToList();

            int countReport = normative.Count;
            int currentIndex = StartPosition;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[2];

            CopyNullCells(ObjWorkSheet, countReport, StartPosition);

            // с июля нормативы другие 
            if (Convert.ToInt32(_yymm) >= 2107)
            {
                foreach (var item in newNormative)
                {
                    ObjWorkSheet.Cells[2, item.Key] = item.Value;

                }
            }

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
                ObjWorkSheet.Cells[currentIndex, 12] = data.Normative.MeeDayHospPlan;
                ObjWorkSheet.Cells[currentIndex, 13] = data.Normative.MeeDayHospTarget;
                ObjWorkSheet.Cells[currentIndex, 15] = data.Normative.BillsHosp;
                ObjWorkSheet.Cells[currentIndex, 16] = data.Normative.MeeHospPlan;
                ObjWorkSheet.Cells[currentIndex, 17] = data.Normative.MeeHospTarget;
                ObjWorkSheet.Cells[currentIndex, 19] = data.Normative.EkmpOutMoPlan;
                ObjWorkSheet.Cells[currentIndex, 20] = data.Normative.EkmpOutMoTarget;
                ObjWorkSheet.Cells[currentIndex, 22] = data.Normative.EkmpAppPlan;
                ObjWorkSheet.Cells[currentIndex, 23] = data.Normative.EkmpAppTarget;
                ObjWorkSheet.Cells[currentIndex, 25] = data.Normative.EkmpDayHospPlan;
                ObjWorkSheet.Cells[currentIndex, 26] = data.Normative.EkmpDayHospTarget;
                ObjWorkSheet.Cells[currentIndex, 28] = data.Normative.EkmpHospPlan;
                ObjWorkSheet.Cells[currentIndex++, 29] = data.Normative.EkmpHospTarget;
            }

           

        }

        private void FillPersonnel(CReportPg[] reports)
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
                ObjWorkSheet.Cells[currentIndex, 3] = data.Personnel.Specialist;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Personnel.MekFullTime;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Personnel.MekRemote;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Personnel.ExpertsFullTime;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Personnel.ExpertsRemote;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Personnel.ExpertsEkmpRegion;
                ObjWorkSheet.Cells[currentIndex, 11] = data.Personnel.ExpertsEkmpRemote;
                ObjWorkSheet.Cells[currentIndex, 12] = data.Personnel.ExpertsEkmpRegionOnko;
                ObjWorkSheet.Cells[currentIndex, 13] = data.Personnel.ExpertsEkmpRemoteOnko;
                ObjWorkSheet.Cells[currentIndex, 14] = data.Personnel.ExpertsEkmpRegister;
                ObjWorkSheet.Cells[currentIndex, 15] = data.Personnel.ExpertsEkmpRegisterRemote;
                ObjWorkSheet.Cells[currentIndex, 16] = data.Personnel.ExpertsEkmpRegisterOnko;
                ObjWorkSheet.Cells[currentIndex, 17] = data.Personnel.ExpertsEkmpRegisterRemoteOnko;
                ObjWorkSheet.Cells[currentIndex, 18] = data.Personnel.ExpertsOmsFullTime;
                ObjWorkSheet.Cells[currentIndex, 19] = data.Personnel.ExpertsOmsRemote;
                ObjWorkSheet.Cells[currentIndex, 20] = data.Personnel.ExpertsOmsEkmpFullTime;
                ObjWorkSheet.Cells[currentIndex++, 21] = data.Personnel.ExpertsOmsEkmpRemote;
            }
        }

        private void FillFinance(CReportPg[] reports)
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

        private void FillExpertises(CReportPg[] reports)
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
                ObjWorkSheet.Cells[currentIndex, 3] = data.Expertise.BillsOnco;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Expertise.BillsVioletion;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Expertise.PaymentBills;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Expertise.PaymentBillsOnco;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Expertise.MeeTarget;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Expertise.MeePlan;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Expertise.CaseMeeTarget;
                ObjWorkSheet.Cells[currentIndex, 11] = data.Expertise.CaseMeePlan;
                ObjWorkSheet.Cells[currentIndex, 13] = data.Expertise.DefectMeeTarget;
                ObjWorkSheet.Cells[currentIndex, 14] = data.Expertise.DefectMeePlan;
                ObjWorkSheet.Cells[currentIndex, 16] = data.Expertise.EkmpTarget;
                ObjWorkSheet.Cells[currentIndex, 17] = data.Expertise.EkmpPlan;
                ObjWorkSheet.Cells[currentIndex, 19] = data.Expertise.ThemeCaseEkmpPlan;
                ObjWorkSheet.Cells[currentIndex, 22] = data.Expertise.CaseEkmpTarget;
                ObjWorkSheet.Cells[currentIndex, 23] = data.Expertise.CaseEkmpPlan;
                ObjWorkSheet.Cells[currentIndex, 25] = data.Expertise.DefectEkmpTarget;
                ObjWorkSheet.Cells[currentIndex++, 26] = data.Expertise.DefectEkmpPlan;
            }
        }
    }
}
