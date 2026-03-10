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
    public class ExcelConsolidateTable6StudentsCreator : ExcelBaseCreator<ConsolidateTable6Students[]>
    {
        private string _yymm;

        public ExcelConsolidateTable6StudentsCreator(
                                          string filename,
                                          string header,
                                          string filialName, string yymm) : base(filename, ExcelForm.consT6Students, header, filialName, false)
        {
            _yymm = yymm;
        }

        protected override void FillReport(ConsolidateTable6Students[] report, ConsolidateTable6Students[] yearReport)
        {
            int countReport = report.Length;
            int currentIndex = 6;

            string shortMonth = _yymm[2] == '0'
    ? _yymm.Substring(3)
    : _yymm.Substring(2);

            ObjWorkSheet.Cells[3, 1] = "Страхование студентов за " + shortMonth + " месяц 20" + _yymm.Substring(0, 2) + " года";
            ObjWorkSheet.Cells[4, 4] = YymmUtils.ConvertYymmToDate(_yymm);


            CopyNullCells(ObjWorkSheet, countReport + 2, 6);
            foreach (var data in report)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = currentIndex - 5;
                ObjWorkSheet.Cells[currentIndex, 2] = data.RegionName;
                ObjWorkSheet.Cells[currentIndex, 3] = data.CountUniversity;
                ObjWorkSheet.Cells[currentIndex, 4] = data.CountCollege;
                ObjWorkSheet.Cells[currentIndex, 5] = data.CountInsured;
                ObjWorkSheet.Cells[currentIndex, 6] = data.Comments;

                currentIndex++;
            }

        // === ИТОГОВАЯ СТРОКА ===
        int totalRow = currentIndex;

        ObjWorkSheet.Cells[totalRow, 1] = "ИТОГО";

            for (int col = 3; col <= 5; col++)
            {
                    // Формула: =SUM(C3:C{lastRow})
                    string colLetter = GetColumnLetter(col);
    ObjWorkSheet.Cells[totalRow, col].Formula = $"=SUM({colLetter}6:{colLetter}{totalRow - 1})";
            }

            // Выделяем итоговую строку жирным
            ObjWorkSheet.Range[$"A{totalRow}:F{totalRow}"].Font.Bold = true;
        }

        // Вспомогательный метод: номер столбца → буква (2→B, 3→C, ..., 102→CV)
        private string GetColumnLetter(int colNum)
{
    int dividend = colNum;
    string columnName = string.Empty;

    while (dividend > 0)
    {
        int modulo = (dividend - 1) % 26;
        columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
        dividend = (dividend - modulo) / 26;
    }

    return columnName;
}
    }
}
