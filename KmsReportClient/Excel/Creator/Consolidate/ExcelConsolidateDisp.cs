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
    class ExcelConsolidateDisp : ExcelBaseCreator<ConsolidateDisp[]>
    {
        public ExcelConsolidateDisp(
           string filename,
           string header,
           string filialName) : base(filename, ExcelForm.disp, header, filialName, false) { }

        protected override void FillReport(ConsolidateDisp[] report, ConsolidateDisp[] yearReport)
        {
            FillEkmp(report);
            FillComplaint(report);
            FillProtection(report);
            FillFinance(report);
            FillMee(report);
            FillMek(report);
        }

        private void FillComplaint(ConsolidateDisp[] reports)
        {
            var complaint = reports.Select(r => new { r.Filial, r.Complaint }).ToList();

            int countReport = complaint.Count;
            int currentIndex = 5;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            CopyNullCells(ObjWorkSheet, countReport, currentIndex);

            int counter = 1;
            foreach (var data in complaint)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Complaint.Row361Gr7;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Complaint.Row365Gr7;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Complaint.Row37Gr7;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Complaint.Row371Gr7;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Complaint.Row372Gr7;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Complaint.Row3721Gr7;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Complaint.Row373Gr7;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Complaint.Row3731Gr7;
                ObjWorkSheet.Cells[currentIndex, 11] = data.Complaint.Row462Gr7;
                ObjWorkSheet.Cells[currentIndex, 12] = data.Complaint.Row47Gr7;
                ObjWorkSheet.Cells[currentIndex, 13] = data.Complaint.Row471Gr7;
                ObjWorkSheet.Cells[currentIndex, 14] = data.Complaint.Row472Gr7;
                ObjWorkSheet.Cells[currentIndex, 15] = data.Complaint.Row4721Gr7;
                ObjWorkSheet.Cells[currentIndex, 16] = data.Complaint.Row473Gr7;
                ObjWorkSheet.Cells[currentIndex, 17] = data.Complaint.Row4731Gr7;
                ObjWorkSheet.Cells[currentIndex++, 18] = data.Complaint.Row49Gr7;


            }

        }

        private void FillProtection(ConsolidateDisp[] reports)
        {
            var protection = reports.Select(r => new { r.Filial, r.Protection }).ToList();

            int countReport = protection.Count;
            int currentIndex = 5;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[2];
            CopyNullCells(ObjWorkSheet, countReport, 5);

            int counter = 1;
            foreach (var data in protection)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Protection.Row161Gr3;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Protection.Row165Gr3;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Protection.Row17Gr3;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Protection.Row171Gr3;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Protection.Row172Gr3;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Protection.Row1721Gr3;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Protection.Row173Gr3;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Protection.Row1731Gr3;

                ObjWorkSheet.Cells[currentIndex, 11] = data.Protection.Row161Gr6;
                ObjWorkSheet.Cells[currentIndex, 12] = data.Protection.Row165Gr6;
                ObjWorkSheet.Cells[currentIndex, 13] = data.Protection.Row17Gr6;
                ObjWorkSheet.Cells[currentIndex, 14] = data.Protection.Row171Gr6;
                ObjWorkSheet.Cells[currentIndex, 15] = data.Protection.Row172Gr6;
                ObjWorkSheet.Cells[currentIndex, 16] = data.Protection.Row1721Gr6;
                ObjWorkSheet.Cells[currentIndex, 17] = data.Protection.Row173Gr6;
                ObjWorkSheet.Cells[currentIndex++, 18] = data.Protection.Row1731Gr6;
            }
        }

        private void FillMek(ConsolidateDisp[] reports)
        {
            var mek = reports.Select(r => new { r.Filial, r.Mek }).ToList();

            int countReport = mek.Count;
            int currentIndex = 5;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[3];
            CopyNullCells(ObjWorkSheet, countReport, currentIndex);

            int counter = 1;
            foreach (var data in mek)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Mek.Row1Gr3;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Mek.Row12Gr3;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Mek.Row4Gr3;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Mek.Row41Gr3;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Mek.Row411Gr3;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Mek.Row42Gr3;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Mek.Row421Gr3;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Mek.Row43Gr3;
                ObjWorkSheet.Cells[currentIndex, 11] = data.Mek.Row431Gr3;
                ObjWorkSheet.Cells[currentIndex, 11] = data.Mek.Row431Gr3;
                ObjWorkSheet.Cells[currentIndex, 12] = data.Mek.Row44Gr3;
                ObjWorkSheet.Cells[currentIndex, 13] = data.Mek.Row441Gr3;
                ObjWorkSheet.Cells[currentIndex, 14] = data.Mek.Row45Gr3;
                ObjWorkSheet.Cells[currentIndex, 15] = data.Mek.Row451Gr3;
                ObjWorkSheet.Cells[currentIndex, 16] = data.Mek.Row5Gr3;
                ObjWorkSheet.Cells[currentIndex, 17] = data.Mek.Row52Gr3;
                ObjWorkSheet.Cells[currentIndex++, 18] = data.Mek.Row52Gr3;
         
            }
        }

        private void FillMee(ConsolidateDisp[] reports)
        {
            var mee = reports.Select(r => new { r.Filial, r.Mee }).ToList();

            int countReport = mee.Count;
            int currentIndex = 6;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[4];
            CopyNullCells(ObjWorkSheet, countReport, currentIndex);

            int counter = 1;
            foreach (var data in mee)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Mee.Row21Gr3;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Mee.Row22Gr3;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Mee.Row221Gr3;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Mee.Row222Gr3;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Mee.Row223Gr3;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Mee.Row24Gr3;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Mee.Row241Gr3;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Mee.Row26Gr10;

                ObjWorkSheet.Cells[currentIndex, 11] = data.Mee.Row531Gr3;
                ObjWorkSheet.Cells[currentIndex, 12] = data.Mee.Row531Gr10;

                ObjWorkSheet.Cells[currentIndex, 13] = data.Mee.Row5311Gr3;
                ObjWorkSheet.Cells[currentIndex, 14] = data.Mee.Row5311Gr10;

                ObjWorkSheet.Cells[currentIndex, 15] = data.Mee.Row54Gr3;
                ObjWorkSheet.Cells[currentIndex, 16] = data.Mee.Row54Gr10;

                ObjWorkSheet.Cells[currentIndex, 17] = data.Mee.Row55Gr3;
                ObjWorkSheet.Cells[currentIndex, 18] = data.Mee.Row55Gr10;

                ObjWorkSheet.Cells[currentIndex, 19] = data.Mee.Row56Gr3;
                ObjWorkSheet.Cells[currentIndex, 20] = data.Mee.Row56Gr10;

                ObjWorkSheet.Cells[currentIndex, 21] = data.Mee.Row561Gr3;
                ObjWorkSheet.Cells[currentIndex++, 22] = data.Mee.Row561Gr10;

            }
        }

        private void FillEkmp(ConsolidateDisp[] reports)
        {
            var ekmp = reports.Select(r => new { r.Filial, r.Ekmp }).ToList();

            int countReport = ekmp.Count;
            int currentIndex = 6;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[5];
            CopyNullCells(ObjWorkSheet, countReport, currentIndex);

            int counter = 1;
            foreach (var data in ekmp)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Ekmp.Row21Gr3;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Ekmp.Row223Gr3;   

                ObjWorkSheet.Cells[currentIndex, 5] = data.Ekmp.Row25Gr3;   
                ObjWorkSheet.Cells[currentIndex, 6] = data.Ekmp.Row25Gr10;

                ObjWorkSheet.Cells[currentIndex, 7] = data.Ekmp.Row251Gr3;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Ekmp.Row251Gr10;

                ObjWorkSheet.Cells[currentIndex, 9] = data.Ekmp.Row611Gr3;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Ekmp.Row611Gr10;

                ObjWorkSheet.Cells[currentIndex, 11] = data.Ekmp.Row6111Gr3;
                ObjWorkSheet.Cells[currentIndex, 12] = data.Ekmp.Row6111Gr10;

                ObjWorkSheet.Cells[currentIndex, 13] = data.Ekmp.Row62Gr3;
                ObjWorkSheet.Cells[currentIndex, 14] = data.Ekmp.Row62Gr10;

                ObjWorkSheet.Cells[currentIndex, 15] = data.Ekmp.Row621Gr3;
                ObjWorkSheet.Cells[currentIndex, 16] = data.Ekmp.Row621Gr10;

                ObjWorkSheet.Cells[currentIndex, 17] = data.Ekmp.Row63Gr3;
                ObjWorkSheet.Cells[currentIndex, 18] = data.Ekmp.Row63Gr10;

                ObjWorkSheet.Cells[currentIndex, 19] = data.Ekmp.Row631Gr3;
                ObjWorkSheet.Cells[currentIndex, 20] = data.Ekmp.Row631Gr10;

                ObjWorkSheet.Cells[currentIndex, 21] = data.Ekmp.Row632Gr3;
                ObjWorkSheet.Cells[currentIndex, 22] = data.Ekmp.Row632Gr10;

                ObjWorkSheet.Cells[currentIndex, 23] = data.Ekmp.Row633Gr3;
                ObjWorkSheet.Cells[currentIndex, 24] = data.Ekmp.Row633Gr10;

                ObjWorkSheet.Cells[currentIndex, 25] = data.Ekmp.Row64Gr3;
                ObjWorkSheet.Cells[currentIndex, 26] = data.Ekmp.Row64Gr10;

                ObjWorkSheet.Cells[currentIndex, 27] = data.Ekmp.Row641Gr3;
                ObjWorkSheet.Cells[currentIndex, 28] = data.Ekmp.Row641Gr10;

                ObjWorkSheet.Cells[currentIndex, 29] = data.Ekmp.Row642Gr3;
                ObjWorkSheet.Cells[currentIndex, 30] = data.Ekmp.Row642Gr10;

                ObjWorkSheet.Cells[currentIndex, 31] = data.Ekmp.Row643Gr3;
                ObjWorkSheet.Cells[currentIndex, 32] = data.Ekmp.Row643Gr10;

                ObjWorkSheet.Cells[currentIndex, 33] = data.Ekmp.Row644Gr3;
                ObjWorkSheet.Cells[currentIndex, 34] = data.Ekmp.Row644Gr10;

                ObjWorkSheet.Cells[currentIndex, 35] = data.Ekmp.Row645Gr3;
                ObjWorkSheet.Cells[currentIndex, 36] = data.Ekmp.Row645Gr10;

                ObjWorkSheet.Cells[currentIndex, 37] = data.Ekmp.Row651Gr3;
                ObjWorkSheet.Cells[currentIndex, 38] = data.Ekmp.Row651Gr10;

                ObjWorkSheet.Cells[currentIndex, 39] = data.Ekmp.Row652Gr3;
                ObjWorkSheet.Cells[currentIndex, 40] = data.Ekmp.Row652Gr10;

                ObjWorkSheet.Cells[currentIndex, 41] = data.Ekmp.Row653Gr3;
                ObjWorkSheet.Cells[currentIndex++, 42] = data.Ekmp.Row653Gr10;


            }
        }

        private void FillFinance(ConsolidateDisp[] reports)
        {
            var finance = reports.Select(r => new { r.Filial, r.Finance }).ToList();

            int countReport = finance.Count;
            int currentIndex = 5;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[6];
            CopyNullCells(ObjWorkSheet, countReport, currentIndex);

            int counter = 1;
            foreach (var data in finance)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;
                ObjWorkSheet.Cells[currentIndex, 3] = data.Finance.Row4Gr3;
                ObjWorkSheet.Cells[currentIndex, 4] = data.Finance.Row41Gr3;
                ObjWorkSheet.Cells[currentIndex, 5] = data.Finance.Row411Gr3;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Finance.Row412Gr3;
                ObjWorkSheet.Cells[currentIndex, 7] = data.Finance.Row413Gr3;
                ObjWorkSheet.Cells[currentIndex, 8] = data.Finance.Row414Gr3;
                ObjWorkSheet.Cells[currentIndex, 9] = data.Finance.Row51Gr3;
                ObjWorkSheet.Cells[currentIndex, 10] = data.Finance.Row511Gr3;
                ObjWorkSheet.Cells[currentIndex, 11] = data.Finance.Row512Gr3;
                ObjWorkSheet.Cells[currentIndex, 12] = data.Finance.Row513Gr3;
                ObjWorkSheet.Cells[currentIndex, 13] = data.Finance.Row514Gr3;
                ObjWorkSheet.Cells[currentIndex, 14] = data.Finance.Row52Gr3;
                ObjWorkSheet.Cells[currentIndex, 15] = data.Finance.Row521Gr3;
                ObjWorkSheet.Cells[currentIndex, 16] = data.Finance.Row522Gr3;
                ObjWorkSheet.Cells[currentIndex, 17] = data.Finance.Row523Gr3;
                ObjWorkSheet.Cells[currentIndex, 18] = data.Finance.Row53Gr3;
                ObjWorkSheet.Cells[currentIndex, 19] = data.Finance.Row531Gr3;
                ObjWorkSheet.Cells[currentIndex, 20] = data.Finance.Row532Gr3;
                ObjWorkSheet.Cells[currentIndex, 21] = data.Finance.Row533Gr3;
                ObjWorkSheet.Cells[currentIndex, 22] = data.Finance.Row54Gr3;
                ObjWorkSheet.Cells[currentIndex, 23] = data.Finance.Row541Gr3;
                ObjWorkSheet.Cells[currentIndex, 24] = data.Finance.Row542Gr3;
                ObjWorkSheet.Cells[currentIndex, 25] = data.Finance.Row543Gr3;
                ObjWorkSheet.Cells[currentIndex, 26] = data.Finance.Row55Gr3;
                ObjWorkSheet.Cells[currentIndex, 27] = data.Finance.Row551Gr3;
                ObjWorkSheet.Cells[currentIndex, 28] = data.Finance.Row552Gr3;
                ObjWorkSheet.Cells[currentIndex, 29] = data.Finance.Row553Gr3;
                ObjWorkSheet.Cells[currentIndex, 30] = data.Finance.Row56Gr3;
                ObjWorkSheet.Cells[currentIndex, 31] = data.Finance.Row561Gr3;
                ObjWorkSheet.Cells[currentIndex, 32] = data.Finance.Row562Gr3;
                ObjWorkSheet.Cells[currentIndex++, 33] = data.Finance.Row563Gr3;
          
            }
        }


    }
}
