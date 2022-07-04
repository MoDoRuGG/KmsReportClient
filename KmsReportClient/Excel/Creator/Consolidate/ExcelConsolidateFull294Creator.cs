using System;
using System.Collections.Generic;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelConsolidateFull294Creator : ExcelBaseCreator<Consolidate294>
    {
        public ExcelConsolidateFull294Creator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.CFull294, header, filialName, false)
        {
        }

        protected override void FillReport(Consolidate294 reportList, Consolidate294 yearReport)
        {
            var filialList = reportList.Disp13List.Select(x => x.Filial).Distinct().OrderBy(x => x).ToList();
            var countFilial = filialList.Count;

            for (int j = 1; j < 9; j++)
            {
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[j];
                CopyBlock(ObjWorkSheet, countFilial);
            }

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            int i = 0;
            foreach (var filial in filialList)
            {
                var disp13 = reportList.Disp13List.Where(x => x.Filial == filial);
                int position = 4 + 14 * i++;
                FillDisp13(disp13, filial, position);
            }

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[2];
            i = 0;
            foreach (var filial in filialList)
            {
                var disp12 = reportList.Disp12List.Where(x => x.Filial == filial);
                int position = 4 + 14 * i++;
                FillDisp13(disp12, filial, position);
            }

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[3];
            i = 0;
            foreach (var filial in filialList)
            {
                var disp2 = reportList.Disp2List.Where(x => x.Filial == filial);
                int position = 4 + 14 * i++;
                FillDisp2(disp2, filial, position);
            }

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[4];
            i = 0;
            foreach (var filial in filialList)
            {
                var dispensaryObservations = reportList.DispensaryObservationList.Where(x => x.Filial == filial);
                int position = 4 + 14 * i++;
                FillDispensaryObservation(dispensaryObservations, filial, position);
            }

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[5];
            i = 0;
            foreach (var filial in filialList)
            {
                var phoneQuestionary = reportList.PhoneQuestionaryList.Where(x => x.Filial == filial);
                int position = 4 + 14 * i++;
                FillPhoneQuestionaries(phoneQuestionary, filial, position, reportList);
            }

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[6];
            i = 0;
            foreach (var filial in filialList)
            {
                var representatives = reportList.InsuranceRepresentativeList.Where(x => x.Filial == filial);
                int position = 4 + 14 * i++;
                FillInsuranceRepresentatives(representatives, filial, position);
            }

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[7];
            i = 0;
            foreach (var filial in filialList)
            {
                var treatments = reportList.TreatmentList.Where(x => x.Filial == filial);
                int position = 4 + 14 * i++;
                FillTreatments(treatments, filial, position);
            }

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[8];
            i = 0;
            foreach (var filial in filialList)
            {
                var efficiencies = reportList.EfficiencyList.Where(x => x.Filial == filial);
                int position = 4 + 14 * i++;
                FillEfficiencies(efficiencies, filial, position, reportList);
            }
        }

        private void FillEfficiencies(
            IEnumerable<Efficiency> dispList,
            string filial,
            int position,
            Consolidate294 reportList)
        {
            var informCountList = from d13 in reportList.Disp13List
                join d12 in reportList.Disp12List
                    on new {d13.Filial, d13.Month} equals new {d12.Filial, d12.Month}
                select new {d13.Filial, d13.Month, Inform13 = d13.CountPeopleInform, Inform12 = d12.CountPeopleInform};

            ObjWorkSheet.Cells[position, 1] = filial;

            foreach (var month in GlobalConst.Months)
            {
                position++;
                var disp = dispList.SingleOrDefault(x =>
                    x.Month.Equals(month, StringComparison.InvariantCultureIgnoreCase));
                if (disp == null)
                    continue;

                var inform = informCountList.SingleOrDefault(x => x.Filial == disp.Filial && x.Month == disp.Month);
                int inform13 = inform?.Inform13 ?? 0;
                int inform12 = inform?.Inform12 ?? 0;

                ObjWorkSheet.Cells[position, 1] = disp.Month;
                ObjWorkSheet.Cells[position, 2] = disp.Disp13Inform;
                ObjWorkSheet.Cells[position, 4] = disp.Disp12Inform;
                ObjWorkSheet.Cells[position, 6] = disp.AimedDisp2;
                ObjWorkSheet.Cells[position, 7] = disp.Disp2;
                ObjWorkSheet.Cells[position, 9] = disp.SubjectToDisp;
                ObjWorkSheet.Cells[position, 10] = inform13;
                ObjWorkSheet.Cells[position, 11] = inform12;
            }
        }

        private void FillTreatments(IEnumerable<Treatment> dispList, string filial, int position)
        {
            ObjWorkSheet.Cells[position, 1] = filial;
            foreach (var month in GlobalConst.Months)
            {
                position++;
                var disp = dispList.SingleOrDefault(x =>
                    x.Month.Equals(month, StringComparison.InvariantCultureIgnoreCase));
                if (disp == null)
                    continue;

                ObjWorkSheet.Cells[position, 1] = disp.Month;
                ObjWorkSheet.Cells[position, 2] = disp.Oral;
                ObjWorkSheet.Cells[position, 3] = disp.OralSecond;
                ObjWorkSheet.Cells[position, 4] = disp.OralThird;
                ObjWorkSheet.Cells[position, 5] = disp.Written;
                ObjWorkSheet.Cells[position, 6] = disp.WrittenSecond;
                ObjWorkSheet.Cells[position, 7] = disp.WrittenThird;
                ObjWorkSheet.Cells[position, 8] = disp.WrittenDoctor;
            }
        }

        private void FillInsuranceRepresentatives(IEnumerable<InsuranceRepresentative> dispList, string filial,
            int position)
        {
            ObjWorkSheet.Cells[position, 1] = filial;
            foreach (var month in GlobalConst.Months)
            {
                position++;
                var disp = dispList.SingleOrDefault(
                    x => x.Month.Equals(month, StringComparison.InvariantCultureIgnoreCase));
                if (disp == null)
                    continue;

                ObjWorkSheet.Cells[position, 1] = disp.Month;
                ObjWorkSheet.Cells[position, 2] = disp.FirstLevel;
                ObjWorkSheet.Cells[position, 3] = disp.FirstLevelTraining;
                ObjWorkSheet.Cells[position, 4] = disp.SecondLevel;
                ObjWorkSheet.Cells[position, 5] = disp.SecondLevelTraining;
                ObjWorkSheet.Cells[position, 6] = disp.ThirdLevel;
                ObjWorkSheet.Cells[position, 7] = disp.ThirdLevelTraining;
            }
        }

        private void FillPhoneQuestionaries(
            IEnumerable<PhoneQuestionary> dispList,
            string filial,
            int position,
            Consolidate294 reportList)
        {
            var dispCountList = from d13 in reportList.Disp13List
                join d12 in reportList.Disp12List
                    on new {d13.Filial, d13.Month} equals new {d12.Filial, d12.Month}
                select new {d13.Filial, d13.Month, SumDisp = d13.CountPeoplelMo + d12.CountPeoplelMo};

            ObjWorkSheet.Cells[position, 1] = filial;

            foreach (var month in GlobalConst.Months)
            {
                position++;
                var disp = dispList.SingleOrDefault(
                    x => x.Month.Equals(month, StringComparison.InvariantCultureIgnoreCase));
                if (disp == null)
                    continue;

                int sumDisp = dispCountList.SingleOrDefault(x => x.Filial == disp.Filial && x.Month == disp.Month)
                                  ?.SumDisp ?? 0;
                ObjWorkSheet.Cells[position, 1] = disp.Month;
                ObjWorkSheet.Cells[position, 2] = disp.Prof;
                ObjWorkSheet.Cells[position, 3] = disp.Disp;
                ObjWorkSheet.Cells[position, 5] = sumDisp;
            }
        }

        private void FillDispensaryObservation(
            IEnumerable<DispensaryObservation> dispList,
            string filial,
            int position)
        {
            ObjWorkSheet.Cells[position, 1] = filial;
            foreach (var month in GlobalConst.Months)
            {
                position++;
                var disp = dispList.SingleOrDefault(
                    x => x.Month.Equals(month, StringComparison.InvariantCultureIgnoreCase));
                if (disp == null)
                    continue;

                ObjWorkSheet.Cells[position, 1] = disp.Month;
                ObjWorkSheet.Cells[position, 2] = disp.CountPeopleInform;
                ObjWorkSheet.Cells[position, 3] = disp.Onco;
                ObjWorkSheet.Cells[position, 4] = disp.Endo;
                ObjWorkSheet.Cells[position, 5] = disp.Broncho;
                ObjWorkSheet.Cells[position, 6] = disp.BloodCirculatory;
                ObjWorkSheet.Cells[position, 7] = disp.NotInfection;
            }
        }

        private void FillDisp2(IEnumerable<Dispanserization> dispList, string filial, int position)
        {
            ObjWorkSheet.Cells[position, 1] = filial;
            foreach (var month in GlobalConst.Months)
            {
                position++;
                var disp = dispList.SingleOrDefault(
                    x => x.Month.Equals(month, StringComparison.InvariantCultureIgnoreCase));
                if (disp == null)
                    continue;

                ObjWorkSheet.Cells[position, 1] = disp.Month;
                ObjWorkSheet.Cells[position, 2] = disp.CountPeoplelMo;
                ObjWorkSheet.Cells[position, 3] = disp.Sms;
                ObjWorkSheet.Cells[position, 4] = disp.Post;
                ObjWorkSheet.Cells[position, 5] = disp.Phone;
                ObjWorkSheet.Cells[position, 6] = disp.Messangers;
                ObjWorkSheet.Cells[position, 7] = disp.EMail;
                ObjWorkSheet.Cells[position, 8] = disp.Address;
                ObjWorkSheet.Cells[position, 9] = disp.AnotherType;
            }
        }

        private void FillDisp13(IEnumerable<Dispanserization> dispList, string filial, int position)
        {
            ObjWorkSheet.Cells[position, 1] = filial;
            foreach (var month in GlobalConst.Months)
            {
                position++;
                var disp = dispList.SingleOrDefault(
                    x => x.Month.Equals(month, StringComparison.InvariantCultureIgnoreCase));
                if (disp == null)
                    continue;

                ObjWorkSheet.Cells[position, 1] = disp.Month;
                ObjWorkSheet.Cells[position, 2] = disp.CountPeoplelMo;
                ObjWorkSheet.Cells[position, 3] = disp.CountPeopleInform;
                ObjWorkSheet.Cells[position, 5] = disp.CountPeopleRepeatInform;
                ObjWorkSheet.Cells[position, 7] = disp.Sms;
                ObjWorkSheet.Cells[position, 8] = disp.Post;
                ObjWorkSheet.Cells[position, 9] = disp.Phone;
                ObjWorkSheet.Cells[position, 10] = disp.Messangers;
                ObjWorkSheet.Cells[position, 11] = disp.EMail;
                ObjWorkSheet.Cells[position, 12] = disp.Address;
                ObjWorkSheet.Cells[position, 13] = disp.AnotherType;
            }
        }

        private void CopyBlock(Worksheet sheet, int count)
        {
            for (int k = 1; k < count; k++)
            {
                var r = sheet.Range["4:17", Type.Missing];
                r.Copy(Type.Missing);
                r = sheet.Range[Convert.ToString(4 + 14 * (k - 1)) + ":" + Convert.ToString(17 + 14 * (k - 1)),
                    Type.Missing];
                r.Insert(XlInsertShiftDirection.xlShiftDown);
            }
        }
    }
}