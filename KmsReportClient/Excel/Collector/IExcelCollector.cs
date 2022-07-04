using KmsReportClient.External;

namespace KmsReportClient.Excel.Collector
{
    interface IExcelCollector
    {
        void Collect(string filename, string form, AbstractReport report);
    }
}
