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
    public class ExcelPVPLoadCreator : ExcelBaseCreator<ReportPVPLoad>
    {

        private EndpointSoap _client;

        public ExcelPVPLoadCreator(
          string filename,
          ExcelForm reportName,
          string header,
          string _regionCode, EndpointSoap client, string regionCode) : base(filename, reportName, header, regionCode, false)
        {
            _client = client;
        }



        protected override void FillReport(ReportPVPLoad report, ReportPVPLoad yearReport)
        {
            var StartPosition = 4;
            int countReport = report.Data.Length;
            int currentIndex = StartPosition;
            CopyNullCells(ObjWorkSheet, countReport+1, StartPosition);
            for (int i = 0; i < countReport; i++)
            {
                var data = report.Data.FirstOrDefault(x => x.RowNumID == i);
                if (data != null) 
                {
                    ObjWorkSheet.Cells[StartPosition + i, 1] = data.RowNumID+1;
                    ObjWorkSheet.Cells[StartPosition + i, 2] = data.PVP_name;
                    ObjWorkSheet.Cells[StartPosition + i, 3] = data.location_of_the_office;
                    ObjWorkSheet.Cells[StartPosition + i, 4] = data.number_of_insured_by_beginning_of_year;
                    ObjWorkSheet.Cells[StartPosition + i, 5] = data.number_of_insured_by_reporting_date;
                    ObjWorkSheet.Cells[StartPosition + i, 6] = data.population_dynamics;
                    ObjWorkSheet.Cells[StartPosition + i, 7] = data.specialist;
                    ObjWorkSheet.Cells[StartPosition + i, 8] = data.conditions_of_employment;
                    ObjWorkSheet.Cells[StartPosition + i, 9] = data.PVP_plan;
                    ObjWorkSheet.Cells[StartPosition + i, 10] = data.registered_total_citizens;
                    ObjWorkSheet.Cells[StartPosition + i, 11] = data.newly_insured;
                    ObjWorkSheet.Cells[StartPosition + i, 12] = data.attracted_by_agents;
                    ObjWorkSheet.Cells[StartPosition + i, 13] = data.issued_by_PEO_and_extracts_from_ERZL;
                    ObjWorkSheet.Cells[StartPosition + i, 16] = data.workload_per_day_for_specialist;
                    ObjWorkSheet.Cells[StartPosition + i, 17] = data.appeals_through_EPGU;
                    ObjWorkSheet.Cells[StartPosition + i, 18] = data.notes;

                }
            }
        }
    }
}
