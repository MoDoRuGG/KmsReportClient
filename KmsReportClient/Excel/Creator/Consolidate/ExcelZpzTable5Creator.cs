using System.Collections.Generic;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelZpzTable5Creator : ExcelBaseCreator<ConsolidateZpzTable5[]>
    {
        public ExcelZpzTable5Creator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.ZpzQ2025_t5, header, filialName, false) { }

        protected override void FillReport(ConsolidateZpzTable5[] reports, ConsolidateZpzTable5[] yearReport)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            int currentIndex = 6;

            foreach (var report in reports.OrderBy(x => x.Filial))
            {
                if (!string.IsNullOrEmpty(report.Filial))
                {
                    ObjWorkSheet.Cells[currentIndex, 1] = report.Filial;

                    // Добавление данных из свойства Data
                    int colIndex = 2; // Предполагается, что первая колонка уже занята названием филиала

                    ObjWorkSheet.Cells[currentIndex, 2] = report.RowNum;
                    ObjWorkSheet.Cells[currentIndex, 3] = report.CountSmo;
                    ObjWorkSheet.Cells[currentIndex, 4] = report.CountSmoAnother;
                    ObjWorkSheet.Cells[currentIndex, 5] = report.CountInsured;
                    ObjWorkSheet.Cells[currentIndex, 6] = report.CountInsuredRepresentative;
                    ObjWorkSheet.Cells[currentIndex, 7] = report.CountTfoms;
                    ObjWorkSheet.Cells[currentIndex, 8] = report.CountProsecutor;
                    ObjWorkSheet.Cells[currentIndex, 9] = report.CountOutOfSmo;

                    currentIndex++;
                }
            }
        }
    }
}