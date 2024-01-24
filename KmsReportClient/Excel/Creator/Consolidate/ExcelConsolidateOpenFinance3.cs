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

    public class ExcelConsolidateOpenFinance3 : ExcelBaseCreator<ConsolidateOpedFinance_3[]>

    {

        public ExcelConsolidateOpenFinance3(
     string filename,
     string header,
     string filialName) : base(filename, ExcelForm.consOpedFinance3, header, filialName, false)
        {

        }


        protected override void FillReport(ConsolidateOpedFinance_3[] report, ConsolidateOpedFinance_3[] yearReport)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            //CopyNullCellsNew(ObjWorkSheet);

            var filials = report.Select(x => x.RegionName).Distinct().OrderBy(x => x);
            int rowIndex = 5;
            int[] continueColumns = { 15, 31, 47, 63, 67 };
            foreach (var filial in filials)
            {
                var monthData = report.Where(x => x.RegionName == filial);
                int columnIndex = 3;
                ObjWorkSheet.Cells[rowIndex, 2] = filial;
                foreach (var md in monthData.OrderBy(x => x.Yymm))
                {


                    #region Определяем номер следующего столбца
                    if (continueColumns.Contains(columnIndex))
                    #endregion
                    {
                        if (columnIndex == 15 || columnIndex == 31 || columnIndex == 47 || columnIndex == 63 || columnIndex == 67)
                    {

                            columnIndex++;
                            columnIndex++;
                            ObjWorkSheet.Cells[rowIndex, ++columnIndex] = md.Notes;
                            columnIndex++;
                        }
                    }
                    
                    else
                    {
                        ObjWorkSheet.Cells[rowIndex, columnIndex++] = md.Fact;
                        ObjWorkSheet.Cells[rowIndex, columnIndex++] = md.PlanO;
                        ObjWorkSheet.Cells[rowIndex, ++columnIndex] = md.Notes;
                        columnIndex++;
                    }
                }

                rowIndex += 1;

            }
        }


        protected void CopyNullCellsNew(Worksheet sheet)
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
    }
}
