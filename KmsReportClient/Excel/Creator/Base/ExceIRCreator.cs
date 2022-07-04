using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Base
{
    public class ExceIRCreator : ExcelBaseCreator<ReportInfrormationResponse>
    {

        private EndpointSoap _client;

        private string _regionCode;
   
        public ExceIRCreator(
          string filename,
          ExcelForm reportName,
          string header,
          string filialName, EndpointSoap client, string regionCode) : base(filename, reportName, header, filialName, false)
        {
            _client = client;
            _regionCode = regionCode;



        }

   

        protected override void FillReport(ReportInfrormationResponse report, ReportInfrormationResponse yearReport)
        {
            int sheet = 1;
            foreach (var theme in report.ReportDataList)
            {
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[sheet];

                ObjWorkSheet.Cells[4, 1] = FilialName;
                ObjWorkSheet.Cells[4, 2] = theme.Data.Plan;
                ObjWorkSheet.Cells[4, 3] = theme.Data.Informed;
                ObjWorkSheet.Cells[4, 5] = theme.Data.CountPast;
                ObjWorkSheet.Cells[4, 6] = theme.Data.CountRegistry;


                var yearThemeData = _client.GetIRYearData(new GetIRYearDataRequest(new GetIRYearDataRequestBody
                {
                    fillial = _regionCode,
                    theme = theme.Theme,
                    yymm = report.Yymm
                })).Body.GetIRYearDataResult;

                if (yearThemeData != null)
                {

                    ObjWorkSheet.Cells[4, 9].Value = yearThemeData.Informed;
                    ObjWorkSheet.Cells[4, 11].Value = yearThemeData.CountPast;
                    ObjWorkSheet.Cells[4, 12].Value = yearThemeData.CountRegistry;
                }

                sheet++;

            }

        }
    }
}
