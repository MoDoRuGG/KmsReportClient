using System;
using KmsReportClient.Excel.Creator.Consolidate;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Support;

namespace KmsReportClient.Report.Consolidate
{
    // todo rewrite all consolidate creators with factory pattern
    class ZpzForWebSiteCreator : IConsolidateReportCreator
    {
        public void CreateReport(
            EndpointSoapClient client,
            int year,
            string periodStart,
            string periodEnd,
            string filename)
        {
            decimal yy = year - 2000;
            int mmEnd = 3 * (Array.IndexOf(GlobalConst.Periods, periodStart) + 1);
            string yymm = $"{yy}{mmEnd}";

            var data = client.CreateZpzForWebSite(yymm);           
            var excel = new ExcelConsZpzWebSite(filename, "", filename);
           // excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(filename);
        }
    }
}
