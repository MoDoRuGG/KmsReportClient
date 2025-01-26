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

    public class ExcelConsolidateOpenFinance1 : ExcelBaseCreator<ConsolidateOpedFinance_1[]>

    {

        public ExcelConsolidateOpenFinance1(
     string filename,
     string header,
     string filialName) : base(filename, ExcelForm.consOpedFinance1, header, filialName, false)
        {

        }


        protected override void FillReport(ConsolidateOpedFinance_1[] report, ConsolidateOpedFinance_1[] yearReport)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            CopyNullCellsP1(ObjWorkSheet);

            var filials = report.Select(x => x.RegionName).Distinct().OrderBy(x => x);
            int rowIndex = 5;

            int[] continueColumns = { 5, 9, 14 };
            foreach (var filial in filials)
            {
                var monthData = report.Where(x => x.RegionName == filial);
                int columnIndex = 3;
                ObjWorkSheet.Cells[rowIndex - 1, 2] = filial;
                foreach (var md in monthData.OrderBy(x => x.Yymm))
                {
                    int currentRowIndex = rowIndex;
                    ObjWorkSheet.Cells[currentRowIndex, columnIndex] = md.PlanO;
                    ObjWorkSheet.Cells[++currentRowIndex, columnIndex] = md.Fact;
                    currentRowIndex += 2;
                    ObjWorkSheet.Cells[currentRowIndex, columnIndex] = md.Mee;
                    ObjWorkSheet.Cells[++currentRowIndex, columnIndex] = md.Ekmp;
                    ObjWorkSheet.Cells[++currentRowIndex, columnIndex] = md.Penalty;
                    currentRowIndex += 11;
                    ObjWorkSheet.Cells[currentRowIndex, columnIndex] = md.CountRegularExpertMee;
                    ObjWorkSheet.Cells[++currentRowIndex, columnIndex] = md.CountRegularExpertEkmp;
                    ObjWorkSheet.Cells[++currentRowIndex, columnIndex] = md.CountFreelanceExpert;
                    ObjWorkSheet.Cells[++currentRowIndex, columnIndex] = md.PaymentFreelanceExpert;
                    ObjWorkSheet.Cells[++currentRowIndex, columnIndex] = md.PenaltyTfoms;

                    #region Определяем номер следующего столбца
                    if (continueColumns.Contains(columnIndex))
                    {
                        if (columnIndex == 5)
                        {
                            columnIndex += 2;
                        }

                        if (columnIndex == 9 || columnIndex == 14)
                        {
                            columnIndex += 3;
                        }
                    }
                    else
                    {
                        columnIndex++;
                    }

                    #endregion
                }

                rowIndex += 22;

            }

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[2];
            CopyNullCellsP2(ObjWorkSheet);

            filials = report.Select(x => x.RegionName).Distinct().OrderBy(x => x);
            rowIndex = 5;

            foreach (var filial in filials)
            {
                var monthData = report.Where(x => x.RegionName == filial);
                int columnIndex = 3;
                ObjWorkSheet.Cells[rowIndex - 1, 2] = filial;
                foreach (var md in monthData.OrderBy(x => x.Yymm))
                {
                    int currentRowIndex = rowIndex;
                    ObjWorkSheet.Cells[currentRowIndex, columnIndex] = md.PlanO;
                    ObjWorkSheet.Cells[++currentRowIndex, columnIndex++] = md.Fact;
                }

                rowIndex += 4;

            }


        }


        protected void CopyNullCellsP1(Worksheet sheet)
        {
            int cntS = 4;
            int cntE = 25;
            for (int k = 1; k <= 40 - 1; k++)
            {
                var row = sheet.Range["B" + cntS + ":U" + cntE, Type.Missing];
                row.Copy(Type.Missing);
                cntS += 22;
                cntE += 22;
                row = sheet.Range["B" + cntS + ":U" + cntE, Type.Missing];
                row.Insert(XlInsertShiftDirection.xlShiftDown);

            }
        }


        protected void CopyNullCellsP2(Worksheet sheet)
        {
            int cntS = 4;
            int cntE = 7;
            for (int k = 1; k <= 40 - 1; k++)
            {
                var row = sheet.Range["B" + cntS + ":U" + cntE, Type.Missing];
                row.Copy(Type.Missing);
                cntS += 4;
                cntE += 4;
                row = sheet.Range["B" + cntS + ":U" + cntE, Type.Missing];
                row.Insert(XlInsertShiftDirection.xlShiftDown);

            }
        }
    }
}
