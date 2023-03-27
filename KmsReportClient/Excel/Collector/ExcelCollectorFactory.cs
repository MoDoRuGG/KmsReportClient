using System;
using KmsReportClient.Global;

namespace KmsReportClient.Excel.Collector
{
    class ExcelCollectorFactory
    {
        private readonly IExcelCollector _f262Table3Collector = new F262Table3ExcelCollector();
        private readonly IExcelCollector _pgExcelCollector = new PgExcelCollector();
        private readonly IExcelCollector _pgQExcelCollector = new PgQExcelCollector();
        private readonly IExcelCollector _zpzExcelCollector = new ZpzExcelCollector();
        private readonly IExcelCollector _zpz10ExcelCollector = new ZpzExcelCollector();

        public IExcelCollector GetExcelCollector(string reportType) =>
            reportType switch {
                ReportGlobalConst.Report262 => _f262Table3Collector,
                ReportGlobalConst.ReportPgQ => _pgQExcelCollector,
                ReportGlobalConst.ReportPg => _pgExcelCollector,
                ReportGlobalConst.ReportZpz => _zpzExcelCollector,
                ReportGlobalConst.ReportZpz10 => _zpz10ExcelCollector,
                _ => throw new Exception("Can't find excelCollector for this reportType")
            };
    }
}
