using System.Collections.Generic;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelFFOMSMonthlyVolCreator : ExcelBaseCreator<FFOMSMonthlyVol>
    {
        public ExcelFFOMSMonthlyVolCreator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.MonthlyVol, header, filialName, true) { }

        protected override void FillReport(FFOMSMonthlyVol report, FFOMSMonthlyVol yearReport)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            FillSKP(report);
            FillSDP(report);
            FillAPP(report);
            FillSMP(report);
        }

        private void FillSKP(FFOMSMonthlyVol report)
        {
            int currentIndex = 9;
            foreach (var data in report.FFOMSMonthlyVol_SKP)
            {
                ObjWorkSheet.Cells[currentIndex, 2] = data.CountSluch;
                ObjWorkSheet.Cells[currentIndex, 3] = data.CountAppliedSluch;
                ObjWorkSheet.Cells[currentIndex, 6] = data.CountAppliedSluch;
                ObjWorkSheet.Cells[currentIndex++, 10] = data.CountAppliedSluch;
            }
        }

        private void FillSDP(FFOMSMonthlyVol report)
        {
            int currentIndex = 23;
            foreach (var data in report.FFOMSMonthlyVol_SDP)
            {
                ObjWorkSheet.Cells[currentIndex, 2] = data.CountSluch;
                ObjWorkSheet.Cells[currentIndex, 3] = data.CountAppliedSluch;
                ObjWorkSheet.Cells[currentIndex, 6] = data.CountAppliedSluch;
                ObjWorkSheet.Cells[currentIndex++, 10] = data.CountAppliedSluch;
            }
        }

        private void FillAPP(FFOMSMonthlyVol report)
        {
            int currentIndex = 37;
            foreach (var data in report.FFOMSMonthlyVol_APP)
            {
                ObjWorkSheet.Cells[currentIndex, 2] = data.CountSluch;
                ObjWorkSheet.Cells[currentIndex, 3] = data.CountAppliedSluch;
                ObjWorkSheet.Cells[currentIndex, 6] = data.CountAppliedSluch;
                ObjWorkSheet.Cells[currentIndex++, 10] = data.CountAppliedSluch;
            }
        }

        private void FillSMP(FFOMSMonthlyVol report)
        {
            int currentIndex = 51;
            foreach (var data in report.FFOMSMonthlyVol_SMP)
            {
                ObjWorkSheet.Cells[currentIndex, 2] = data.CountSluch;
                ObjWorkSheet.Cells[currentIndex, 3] = data.CountAppliedSluch;
                ObjWorkSheet.Cells[currentIndex, 6] = data.CountAppliedSluch;
                ObjWorkSheet.Cells[currentIndex++, 10] = data.CountAppliedSluch;
            }
        }
    }
}
