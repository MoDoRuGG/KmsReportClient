using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    public class ExcelConsolidateZpzTable10Creator : ExcelBaseCreator<ConsolidateZpzTable10[]>
    {
        private string _yymm;

        public ExcelConsolidateZpzTable10Creator(
                                          string filename,
                                          string header,
                                          string filialName, string yymm) : base(filename, ExcelForm.Zpz10Cons, header, filialName, false)
        {
            _yymm = yymm;
        }

        protected override void FillReport(ConsolidateZpzTable10[] report, ConsolidateZpzTable10[] yearReport)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            var columnIndex = 3;
            for (int i = 7; i <= 107; i++)
            {
                string rowNum = ObjWorkSheet.Cells[i, 2].Text;
                if (!string.IsNullOrEmpty(rowNum))
                {
                    var rowData = report?.SingleOrDefault(x => x.RowNum == rowNum);
                    if (rowData != null)
                    {
                        if (rowData.RowNum.StartsWith("8"))
                        {
                            ObjWorkSheet.Cells[i, columnIndex + 1] = rowData.ByMonth;
                        }
                        else if (rowData.RowNum == "7.5")
                        {
                            ObjWorkSheet.Cells[i, columnIndex] = rowData.Yearly;
                            ObjWorkSheet.Cells[i, columnIndex + 1] = "X";
                        }
                        else
                        {
                            ObjWorkSheet.Cells[i, columnIndex] = rowData.Yearly;
                            ObjWorkSheet.Cells[i, columnIndex + 1] = rowData.ByMonth;
                        }
                    }
                }
            }
        }
    }
}
