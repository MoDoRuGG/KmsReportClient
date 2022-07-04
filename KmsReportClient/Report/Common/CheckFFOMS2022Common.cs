using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KmsReportClient.External;

namespace KmsReportClient.Report.Common
{
    public class CheckFFOMS2022Common
    {
        private readonly EndpointSoap _client;
        public CheckFFOMS2022Common(EndpointSoap client)
        {
            _client = client;
        }

        public CheckFFOMS2022CommonData GetFFOMS2022CommonData(string year, string idRegion)
        {
            var response = _client.GetCheckFFOMS2022CommonData(new GetCheckFFOMS2022CommonDataRequest
            {
                Body = new GetCheckFFOMS2022CommonDataRequestBody
                {
                    year = year,
                    idReport = idRegion
                }
            }).Body.GetCheckFFOMS2022CommonDataResult;

            return response;
        }
    }
}
