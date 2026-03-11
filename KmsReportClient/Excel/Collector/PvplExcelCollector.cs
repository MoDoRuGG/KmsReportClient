using System;
using System.Collections.Generic;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Model;
using KmsReportClient.Support;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Collector
{
    class PvplExcelCollector : ExcelBaseCollector
    {
        protected override void FillReport(string form, AbstractReport destReport, AbstractReport srcReport)
        {
            var destData = destReport as ReportPVPLoad ??
                          throw new Exception("Can't cast destReport to ReportPVPLoad");
            var srcData = srcReport as ReportPVPLoad ??
                         throw new Exception("Can't cast srcReport to ReportPVPLoad");

            destData.Data = srcData.Data;
        }

        protected override AbstractReport CollectReportData(string form)
        {
            var dataList = new List<PVPload>();
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1]; // Первый лист

            // Находим последнюю заполненную строку
            int lastRow = GetLastRow();
            int startRow = 4; // Данные начинаются с 4-й строки (после заголовков)

            // Проверяем, есть ли данные
            if (lastRow < startRow)
                return new ReportPVPLoad { Data = Array.Empty<PVPload>() };

            for (int i = startRow; i <= lastRow; i++)
            {
                // Проверяем, не строка ли это "Всего по филиалу"
                var cellA = ObjWorkSheet.Cells[i, 1].Text?.ToString()?.Trim();
                if (cellA == "Всего по филиалу:" || string.IsNullOrEmpty(cellA))
                    continue;

                var data = new PVPload
                {
                    RowNumID = dataList.Count,
                    PVP_name = GetCellText(2, i),
                    location_of_the_office = GetCellText(3, i),
                    number_of_insured_by_beginning_of_year = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 4].Text),
                    number_of_insured_by_reporting_date = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 5].Text),
                    population_dynamics = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 6].Text),
                    specialist = GetCellText(7, i),
                    conditions_of_employment = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, 8].Text),
                    PVP_plan = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 9].Text),
                    registered_total_citizens = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 10].Text),
                    newly_insured = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 11].Text),
                    attracted_by_agents = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 12].Text),
                    issued_by_PEO_and_extracts_from_ERZL = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 13].Text),
                    // Колонки 14 и 15 (Всего обслужено и отклонения) - расчетные, не читаем
                    workload_per_day_for_specialist = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, 16].Text),
                    appeals_through_EPGU = GlobalUtils.TryParseInt(ObjWorkSheet.Cells[i, 17].Text),
                    notes = GetCellText(18, i)
                };

                dataList.Add(data);
            }

            return new ReportPVPLoad { Data = dataList.ToArray() };
        }

        private string GetCellText(int column, int row)
        {
            var value = ObjWorkSheet.Cells[row, column].Text?.ToString();
            return string.IsNullOrEmpty(value) ? "" : value.Trim();
        }
    }
}