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
    public class ExceCadreCreator : ExcelBaseCreator<ReportCadre>
    {

        private EndpointSoap _client;

        private string _regionCode;

        public ExceCadreCreator(
          string filename,
          ExcelForm reportName,
          string header,
          string filialName, EndpointSoap client, string regionCode) : base(filename, reportName, header, filialName, false)
        {
            _client = client;
            _regionCode = regionCode;



        }



        protected override void FillReport(ReportCadre report, ReportCadre yearReport)
        {
            //if (report.IdFlow != 0)
            //{
                int sheet = 1;

                foreach (var theme in report.ReportDataList)
                {
                    ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[sheet];

                    ObjWorkSheet.Cells[7, 2] = FilialName;
                    ObjWorkSheet.Cells[7, 3] = theme.Data.count_itog_state;
                    ObjWorkSheet.Cells[7, 4] = theme.Data.count_itog_fact;
                    ObjWorkSheet.Cells[7, 5] = theme.Data.count_itog_vacancy;
                    ObjWorkSheet.Cells[7, 6] = theme.Data.count_leader_state;
                    ObjWorkSheet.Cells[7, 7] = theme.Data.count_leader_fact;
                    ObjWorkSheet.Cells[7, 8] = theme.Data.count_leader_vacancy;
                    ObjWorkSheet.Cells[7, 9] = theme.Data.count_deputy_leader_state;
                    ObjWorkSheet.Cells[7, 10] = theme.Data.count_deputy_leader_fact;
                    ObjWorkSheet.Cells[7, 11] = theme.Data.count_deputy_leader_vacancy;
                    ObjWorkSheet.Cells[7, 12] = theme.Data.count_expert_doctor_state;
                    ObjWorkSheet.Cells[7, 13] = theme.Data.count_expert_doctor_fact;
                    ObjWorkSheet.Cells[7, 14] = theme.Data.count_expert_doctor_vacancy;
                    ObjWorkSheet.Cells[7, 15] = theme.Data.count_grf15;
                    ObjWorkSheet.Cells[7, 16] = theme.Data.count_grf16;
                    ObjWorkSheet.Cells[7, 17] = theme.Data.count_grf17;
                    ObjWorkSheet.Cells[7, 18] = theme.Data.count_grf18;
                    ObjWorkSheet.Cells[7, 19] = theme.Data.count_grf19;
                    ObjWorkSheet.Cells[7, 20] = theme.Data.count_grf20;
                    ObjWorkSheet.Cells[7, 21] = theme.Data.count_grf21;
                    ObjWorkSheet.Cells[7, 22] = theme.Data.count_grf22;
                    ObjWorkSheet.Cells[7, 23] = theme.Data.count_grf23;
                    ObjWorkSheet.Cells[7, 24] = theme.Data.count_grf24;
                    ObjWorkSheet.Cells[7, 25] = theme.Data.count_grf25;
                    ObjWorkSheet.Cells[7, 26] = theme.Data.count_grf26;
                    ObjWorkSheet.Cells[7, 27] = theme.Data.count_specialist_state;
                    ObjWorkSheet.Cells[7, 28] = theme.Data.count_specialist_fact;
                    ObjWorkSheet.Cells[7, 29] = theme.Data.count_specialist_vacancy;






                    //var yearThemeData = _client.GetCadreYearData(new GetCadreYearDataRequest(new GetIRYearDataRequestBody
                    //{
                    //    fillial = _regionCode,
                    //    theme = theme.Theme,
                    //    yymm = report.Yymm
                    //})).Body.GetIRYearDataResult;

                    //if (yearThemeData != null)
                    //{

                    //    ObjWorkSheet.Cells[4, 9].Value = yearThemeData.Informed;
                    //    ObjWorkSheet.Cells[4, 11].Value = yearThemeData.CountPast;
                    //    ObjWorkSheet.Cells[4, 12].Value = yearThemeData.CountRegistry;
                    //}

                    sheet++;
                }
            //}
            //else
            //{
            //    return; 
            //}

        }
    }
}
