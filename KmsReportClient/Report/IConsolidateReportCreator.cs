using KmsReportClient.External;

namespace KmsReportClient.Report
{
    interface IConsolidateReportCreator
    {
        void CreateReport(EndpointSoapClient client, int year, string periodStart, string periodEnd, string filename);
    }
}
