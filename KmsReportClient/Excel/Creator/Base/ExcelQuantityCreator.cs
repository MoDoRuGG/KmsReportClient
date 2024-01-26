using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Base
{
    public class ExcelQuantityCreator : ExcelBaseCreator<ReportQuantity>
    {

        private EndpointSoap _client;

        private string _regionCode;

        public ExcelQuantityCreator(
          string filename,
          ExcelForm reportName,
          string header,
          string filialName, EndpointSoap client, string regionCode) : base(filename, reportName, header, filialName, false)
        {
            _client = client;
            _regionCode = regionCode;
        }



        protected override void FillReport(ReportQuantity report, ReportQuantity yearReport)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            if (report != null)
            {
                ObjWorkSheet.Cells[6, 1] = report.Col_1;
                ObjWorkSheet.Cells[6, 2] = report.Col_2;
                ObjWorkSheet.Cells[6, 3] = report.Col_3;
                ObjWorkSheet.Cells[6, 4] = report.Col_4;
                ObjWorkSheet.Cells[6, 5] = report.Col_5;
                ObjWorkSheet.Cells[6, 6] = report.Col_6;
                ObjWorkSheet.Cells[6, 7] = report.Col_7;
                ObjWorkSheet.Cells[6, 8] = report.Col_8;
                ObjWorkSheet.Cells[6, 9] = report.Col_9;
                ObjWorkSheet.Cells[6, 10] = report.Col_10;
                ObjWorkSheet.Cells[6, 11] = report.Col_11;
                ObjWorkSheet.Cells[6, 12] = report.Col_12;
                ObjWorkSheet.Cells[6, 13] = report.Col_13;
                ObjWorkSheet.Cells[6, 14] = report.Col_14;
                ObjWorkSheet.Cells[6, 15] = report.Col_15;
                ObjWorkSheet.Cells[6, 16] = report.Col_16;
            }
        }
    }
}
