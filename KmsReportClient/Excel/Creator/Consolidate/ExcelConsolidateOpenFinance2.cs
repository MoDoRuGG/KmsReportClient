using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    public class ExcelConsolidateOpenFinance2 : ExcelBaseCreator<ConsolidateOpedFinance_2[]>
    {
        public ExcelConsolidateOpenFinance2(
   string filename,
   string header,
   string filialName) : base(filename, ExcelForm.consOpedFinance2, header, filialName, false)
        {
            Debug.WriteLine("asdasd");
        }

        protected override void FillReport(ConsolidateOpedFinance_2[] report, ConsolidateOpedFinance_2[] yearReport)
        {
            try
            {

                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];

                int rowIndex = 4;
                int counter = 1;
                CopyNullCells(ObjWorkSheet, report.Length, rowIndex);
                foreach(var row in report)
                {
                    ObjWorkSheet.Cells[rowIndex, 1] = counter++;
                    ObjWorkSheet.Cells[rowIndex, 2] = row.RegionName;

                    ObjWorkSheet.Cells[rowIndex, 3] = row.Fact1;
                    ObjWorkSheet.Cells[rowIndex, 4] = row.Plan1;

                    ObjWorkSheet.Cells[rowIndex, 6] = row.Fact2;
                    ObjWorkSheet.Cells[rowIndex, 7] = row.Plan2;

                    ObjWorkSheet.Cells[rowIndex, 12] = row.Fact3;
                    ObjWorkSheet.Cells[rowIndex, 13] = row.Plan3;

                    ObjWorkSheet.Cells[rowIndex, 18] = row.Fact4;
                    ObjWorkSheet.Cells[rowIndex, 19] = row.Plan4;

                    ObjWorkSheet.Cells[rowIndex, 24] = row.Fact5;
                    ObjWorkSheet.Cells[rowIndex, 25] = row.Plan5;

                    ObjWorkSheet.Cells[rowIndex, 30] = row.Fact6;
                    ObjWorkSheet.Cells[rowIndex, 31] = row.Plan6;

                    ObjWorkSheet.Cells[rowIndex, 40] = row.Fact7;
                    ObjWorkSheet.Cells[rowIndex, 41] = row.Plan7;

                    ObjWorkSheet.Cells[rowIndex, 46] = row.Fact8;
                    ObjWorkSheet.Cells[rowIndex, 47] = row.Plan8;

                    ObjWorkSheet.Cells[rowIndex, 52] = row.Fact9;
                    ObjWorkSheet.Cells[rowIndex, 53] = row.Plan9;

                    ObjWorkSheet.Cells[rowIndex, 62] = row.Fact10;
                    ObjWorkSheet.Cells[rowIndex, 63] = row.Plan10;

                    ObjWorkSheet.Cells[rowIndex, 68] = row.Fact11;
                    ObjWorkSheet.Cells[rowIndex, 69] = row.Plan11;

                    ObjWorkSheet.Cells[rowIndex, 74] = row.Fact12;
                    ObjWorkSheet.Cells[rowIndex, 75] = row.Plan12;
                    rowIndex++;

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка формирования сводного отчёта:{Environment.NewLine}{Environment.NewLine}{ex.Message}");
            }
        }
    }
}
