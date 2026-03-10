using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using KmsReportClient.Support;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.OLE.Interop;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    public class ExcelConsolidateTable7OldPolisCreator : ExcelBaseCreator<ConsolidateTable7OldPolis[]>
    {
        private string _yymm;

        public ExcelConsolidateTable7OldPolisCreator(
                                          string filename,
                                          string header,
                                          string filialName, string yymm) : base(filename, ExcelForm.consT7OldPolis, header, filialName, false)
        {
            _yymm = yymm;
        }

        protected override void FillReport(ConsolidateTable7OldPolis[] report, ConsolidateTable7OldPolis[] yearReport)
        {
            int countReport = report.Length;
            int currentIndex = 3;

            string shortMonth = _yymm[2] == '0'
    ? _yymm.Substring(3)
    : _yymm.Substring(2);

            ObjWorkSheet.Cells[1, 1] = "Динамика количества полисов ОМС старого образца за " + shortMonth + " месяцев 20" + _yymm.Substring(0, 2) + " года";

            //CopyNullCells(ObjWorkSheet, countReport + 2, 6);
            foreach (var data in report)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = data.RegionName;
                ObjWorkSheet.Cells[currentIndex, 2] = data.CountConstant2019;
                ObjWorkSheet.Cells[currentIndex, 3] = data.CountYearStart;
                ObjWorkSheet.Cells[currentIndex, 4] = data.CurrentQuantity;
                ObjWorkSheet.Cells[currentIndex, 5] = data.CountOldPolis;
                ObjWorkSheet.Cells[currentIndex, 6] = data.ShareFromQuantity;
                ObjWorkSheet.Cells[currentIndex, 7] = data.YearlyDynamic;
                ObjWorkSheet.Cells[currentIndex, 8] = data.From2019Dynamic;

                currentIndex++;
            }

            //        // === ИТОГОВАЯ СТРОКА ===
            //        int totalRow = currentIndex;

            //        ObjWorkSheet.Cells[totalRow, 1] = "ИТОГО";

            //            for (int col = 2; col <= 5; col++)
            //            {
            //                    // Формула: =SUM(C3:C{lastRow})
            //                    string colLetter = GetColumnLetter(col);
            //    ObjWorkSheet.Cells[totalRow, col].Formula = $"=SUM({colLetter}3:{colLetter}{totalRow - 1})";
            //            }

            //            // Выделяем итоговую строку жирным
            //            ObjWorkSheet.Range[$"A{totalRow}:F{totalRow}"].Font.Bold = true;
            //        }

            //        // Вспомогательный метод: номер столбца → буква (2→B, 3→C, ..., 102→CV)
            //        private string GetColumnLetter(int colNum)
            //{
            //    int dividend = colNum;
            //    string columnName = string.Empty;

            //    while (dividend > 0)
            //    {
            //        int modulo = (dividend - 1) % 26;
            //        columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
            //        dividend = (dividend - modulo) / 26;
            //    }

            //    return columnName;
        }
    }
}
