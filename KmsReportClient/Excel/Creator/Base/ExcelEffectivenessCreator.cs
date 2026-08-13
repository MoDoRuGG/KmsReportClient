using System;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Model;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Base
{
    public class ExcelEffectivenessCreator : ExcelBaseCreator<ReportEffectiveness>
    {
        private readonly string _form; // текущая тема из ComboBox

        public ExcelEffectivenessCreator(
            string filename,
            ExcelForm reportName,
            string header,
            string regionCode,
            EndpointSoap client,
            string filialCode,
            string form)
            : base(filename, reportName, header, regionCode, false)
        {
            _form = form;
        }

        protected override void FillReport(ReportEffectiveness report, ReportEffectiveness yearReport)
        {
            const int StartPosition = 4; // первая строка данных в шаблоне

            var theme = report.ReportDataList.SingleOrDefault(x => x.Theme == _form);
            if (theme?.Data == null || theme.Data.Length == 0)
            {
                return;
            }

            int countReport = theme.Data.Length;
            CopyNullCells(ObjWorkSheet, countReport + 1, StartPosition);

            for (int i = 0; i < countReport; i++)
            {
                var data = theme.Data[i];
                if (data == null) continue;

                int row = StartPosition + i;
                ObjWorkSheet.Cells[row, 2] = data.full_name;           // ФИО врача-эксперта
                ObjWorkSheet.Cells[row, 3] = data.expert_busyness;     // Занятость ставки
                ObjWorkSheet.Cells[row, 4] = data.expert_speciality;   // Специальность
                ObjWorkSheet.Cells[row, 5] = data.expertise_type;      // Вид экспертизы
                ObjWorkSheet.Cells[row, 6] = data.mee_quantity_plan;   // МЭЭ
                ObjWorkSheet.Cells[row, 7] = data.mee_quantity_fact;
                ObjWorkSheet.Cells[row, 8] = data.mee_quantity_percent;
                ObjWorkSheet.Cells[row, 9] = data.mee_yeild_plan;
                ObjWorkSheet.Cells[row, 10] = data.mee_yeild_fact;
                ObjWorkSheet.Cells[row, 11] = data.mee_yeild_percent;
                ObjWorkSheet.Cells[row, 12] = data.ekmp_quantity_plan;  // ЭКМП
                ObjWorkSheet.Cells[row, 13] = data.ekmp_quantity_fact;
                ObjWorkSheet.Cells[row, 14] = data.ekmp_quantity_percent;
                ObjWorkSheet.Cells[row, 15] = data.ekmp_yeild_plan;
                ObjWorkSheet.Cells[row, 16] = data.ekmp_yeild_fact;
                ObjWorkSheet.Cells[row, 17] = data.ekmp_yeild_percent;
            }
        }
    }
}