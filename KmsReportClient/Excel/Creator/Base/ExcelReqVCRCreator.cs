using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;
using Org.BouncyCastle.Ocsp;

namespace KmsReportClient.Excel.Creator.Base
{
    public class ExcelReqVCRCreator : ExcelBaseCreator<ReportReqVCR>
    {

        private EndpointSoap _client;

        public ExcelReqVCRCreator(
          string filename,
          ExcelForm reportName,
          string header,
          string _regionCode, EndpointSoap client, string regionCode) : base(filename, reportName, header, regionCode, false)
        {
            _client = client;
        }



        protected override void FillReport(ReportReqVCR report, ReportReqVCR yearReport)
        {
            var StartPosition = 4;
            int countReport = report.ReportDataList.Length;
            int currentIndex = StartPosition;
            CopyNullCells(ObjWorkSheet, countReport+1, StartPosition);
            for (int i = 0; i < countReport; i++)
            {
                var data = report.ReportDataList.FirstOrDefault(x => x.Data. == i);
                if (data != null) 
                {
                    ObjWorkSheet.Cells[StartPosition + i, 1] = data.FIO;
                    ObjWorkSheet.Cells[StartPosition + i, 2] = data.Speciality;
                    ObjWorkSheet.Cells[StartPosition + i, 3] = data.Period;
                    ObjWorkSheet.Cells[StartPosition + i, 4] = data.CountEKMP;
                    ObjWorkSheet.Cells[StartPosition + i, 5] = data.AmountSank;
                    ObjWorkSheet.Cells[StartPosition + i, 6] = data.AmountPayment;
                    ObjWorkSheet.Cells[StartPosition + i, 7] = data.ProvidedBy;
                    ObjWorkSheet.Cells[StartPosition + i, 8] = data.Comments;

                }
            }
        }
    }
}
