using System.Collections.Generic;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelFFOMSLethalEKMPCreator : ExcelBaseCreator<FFOMSLethalEKMP[]>
    {
        public ExcelFFOMSLethalEKMPCreator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.FFOMSLethalEKMP, header, filialName, false) { }

        protected override void FillReport(FFOMSLethalEKMP[] reports, FFOMSLethalEKMP[] yearReport)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            int currentIndex = 2;

            foreach (var report in reports.OrderBy(x => x.Filial))
            {
                if (!string.IsNullOrEmpty(report.Filial))
                {
                    ObjWorkSheet.Cells[currentIndex, 1] = report.Filial;

                    // Добавление данных из свойства Data
                    int colIndex = 2; // Предполагается, что первая колонка уже занята названием филиала

                        ObjWorkSheet.Cells[currentIndex, 2] = report.Row1;
                        ObjWorkSheet.Cells[currentIndex, 3] = report.Row12;
                        ObjWorkSheet.Cells[currentIndex, 5] = report.Row121;
                        ObjWorkSheet.Cells[currentIndex, 7] = report.Row11;

                    currentIndex++;
                }
            }
        }
    }
}