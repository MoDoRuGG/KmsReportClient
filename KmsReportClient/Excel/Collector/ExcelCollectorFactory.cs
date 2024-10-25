using System;
using KmsReportClient.Global;

namespace KmsReportClient.Excel.Collector
{
    class ExcelCollectorFactory
    {
        private readonly IExcelCollector _f262Table3Collector = new F262Table3ExcelCollector();
        private readonly IExcelCollector _pgExcelCollector = new PgExcelCollector();
        private readonly IExcelCollector _pgQExcelCollector = new PgQExcelCollector();
        private readonly IExcelCollector _zpzQExcelCollector = new ZpzQExcelCollector();
        private readonly IExcelCollector _zpzExcelCollector = new ZpzExcelCollector();
        private readonly IExcelCollector _zpz10ExcelCollector = new ZpzExcelCollector();
        private readonly IExcelCollector _zpzLethalExcelCollector = new ZpzQExcelCollector();
        private readonly IExcelCollector _zpzQ2025ExcelCollector = new ZpzQ2025ExcelCollector();
        private readonly IExcelCollector _zpz2025ExcelCollector = new Zpz2025ExcelCollector();
        private readonly IExcelCollector _zpz10_2025ExcelCollector = new Zpz2025ExcelCollector();
        private readonly IExcelCollector _zpz2025LethalExcelCollector = new ZpzQ2025ExcelCollector();

        public IExcelCollector GetExcelCollector(string reportType) =>
            reportType switch {
                ReportGlobalConst.Report262 => _f262Table3Collector,
                ReportGlobalConst.ReportPgQ => _pgQExcelCollector,
                ReportGlobalConst.ReportPg => _pgExcelCollector,
                ReportGlobalConst.ReportZpz => _zpzExcelCollector,
                ReportGlobalConst.ReportZpz10 => _zpz10ExcelCollector,
                ReportGlobalConst.ReportZpzQ => _zpzQExcelCollector,
                ReportGlobalConst.ReportZpzLethal => _zpzLethalExcelCollector,
                ReportGlobalConst.ReportZpz2025 => _zpz2025ExcelCollector,
                ReportGlobalConst.ReportZpz10_2025 => _zpz10_2025ExcelCollector,
                ReportGlobalConst.ReportZpzQ2025 => _zpzQ2025ExcelCollector,
                ReportGlobalConst.ReportZpz2025Lethal => _zpz2025LethalExcelCollector,
                _ => throw new Exception("Can't find excelCollector for this reportType")
            };
    }
}
