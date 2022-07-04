using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using KmsReportClient.Support;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Base
{
    public class ExcelVacCreator : ExcelBaseCreator<ReportVaccination>
    {
        private string _regionCode;
        private EndpointSoap _client;
        public ExcelVacCreator(
        string filename,
        ExcelForm reportName,
        string header,
        string filialName, EndpointSoap client, string regionCode) : base(filename, reportName, header, filialName, false)
        {
            _client = client;
            _regionCode = regionCode;

        }


        protected override void FillReport(ReportVaccination report, ReportVaccination yearReport)
        {

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];


            ObjWorkSheet.Cells[3, 2] = YymmUtils.GetMonth(report.Yymm.Substring(2));
            ObjWorkSheet.Cells[3, 15] = YymmUtils.ConvertYymmToDate(report.Yymm).Year;

            ObjWorkSheet.Cells[7, 1] = FilialName;


            ObjWorkSheet.Cells[7, 4] = report.M18_39;

            ObjWorkSheet.Cells[7, 4] = report.M18_39;
            ObjWorkSheet.Cells[7, 5] = report.M40_59;
            ObjWorkSheet.Cells[7, 6] = report.M60_65;
            ObjWorkSheet.Cells[7, 7] = report.M66_74;
            ObjWorkSheet.Cells[7, 8] = report.M75_More;

            ObjWorkSheet.Cells[7, 10] = report.W18_39;
            ObjWorkSheet.Cells[7, 11] = report.W40_54;
            ObjWorkSheet.Cells[7, 12] = report.W55_65;
            ObjWorkSheet.Cells[7, 13] = report.W66_74;
            ObjWorkSheet.Cells[7, 14] = report.W75_More;


            var yearThemeData = _client.GetVacYearData(new GetVacYearDataRequest(new GetVacYearDataRequestBody
            {
                fillial = _regionCode,
                yymm = report.Yymm
            })).Body.GetVacYearDataResult;

            if (yearThemeData != null)
            {
                ObjWorkSheet.Cells[7, 17] = yearThemeData.M18_39;
                ObjWorkSheet.Cells[7, 18] = yearThemeData.M40_59;
                ObjWorkSheet.Cells[7, 19] = yearThemeData.M60_65;
                ObjWorkSheet.Cells[7, 20] = yearThemeData.M66_74;
                ObjWorkSheet.Cells[7, 21] = yearThemeData.M75_More;


                ObjWorkSheet.Cells[7, 23] = yearThemeData.W18_39;
                ObjWorkSheet.Cells[7, 24] = yearThemeData.W40_54;
                ObjWorkSheet.Cells[7, 25] = yearThemeData.W55_65;
                ObjWorkSheet.Cells[7, 26] = yearThemeData.W66_74;
                ObjWorkSheet.Cells[7, 27] = yearThemeData.W75_More;
            }


        }
    }
}
