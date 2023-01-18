using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Model.Enums;
using KmsReportClient.Support;

namespace KmsReportClient.Excel.Creator.Base
{
    public class ExcelConcolidateOpedUCreator : ExcelBaseCreator<ReportOpedU>
    {


        List<string> _notPrintRow = new List<string> { "1.3", "2.3", "3.3" };


        public ExcelConcolidateOpedUCreator(
         string filename,
         ExcelForm reportName,
         string header,
         string filialName) : base(filename, reportName, header, filialName, false) { }


        protected override void FillReport(ReportOpedU report, ReportOpedU yearReport)
        {
            int i = 11;
            for ()
            for (i; i <= 19; i++)
            {
                var exRowNum = Convert.ToString(ObjWorkSheet.Cells[i, 2].Value);
                var rowData = report.ReportDataList.SingleOrDefault(x => x.RowNum == exRowNum);
                if (rowData != null && !_notPrintRow.Contains(exRowNum))
                {
                    ObjWorkSheet.Cells[i, 4] = rowData.App;
                    ObjWorkSheet.Cells[i, 5] = rowData.Ks;
                    ObjWorkSheet.Cells[i, 6] = rowData.Ds;
                    ObjWorkSheet.Cells[i, 7] = rowData.Smp;
                    ObjWorkSheet.Cells[i, 8] = rowData.Notes;
                }

            }

        }
    }
}
