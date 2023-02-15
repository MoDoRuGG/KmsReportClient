using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using KmsReportClient.Support;

namespace KmsReportClient.Excel.Creator.Base
{
    public class ExcelMVCRCreator : ExcelBaseCreator<ReportMonitoringVCR>
    {
        Dictionary<string, DataGridViewRow> _rows;

        public ExcelMVCRCreator(
          string filename,
          ExcelForm reportName,
          string header,
          string filialName, Dictionary<string, DataGridViewRow> rows) : base(filename, reportName, header, filialName, false)
        {
            _rows = rows;
        }

        protected override void FillReport(ReportMonitoringVCR report, ReportMonitoringVCR yearReport)
        {
            ObjWorkSheet.Cells[8, 5] = YymmUtils.GetMonth(report.Yymm.Substring(2, 2)) + " 20" + report.Yymm.Substring(0, 2);
            ObjWorkSheet.Cells[9, 4] = FilialName;

            for (int i = 14; i <= 33; i++)
            {
                string rowNum = Convert.ToString(ObjWorkSheet.Cells[i, 1].Value);
                var data = _rows.FirstOrDefault(x => x.Key == rowNum);

                ObjWorkSheet.Cells[i, 3] = data.Value.Cells[2].Value;
                ObjWorkSheet.Cells[i, 4] = data.Value.Cells[3].Value;
                ObjWorkSheet.Cells[i, 5] = data.Value.Cells[4].Value;

            }

        }
    }
}
