using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Base
{
    public class ExcelProposalCreator : ExcelBaseCreator<ReportProposal>
    {
        private string _yymm;
        public ExcelProposalCreator(
      string filename,
      ExcelForm reportName,
      string header,
      string filialName, string yymm) : base(filename, reportName, header, filialName, false)
        {
            _yymm = yymm;

        }

        protected override void FillReport(ReportProposal report, ReportProposal yearReport)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];

            ObjWorkSheet.Cells[1, 2] = FilialName;
            ObjWorkSheet.Cells[2, 2] = _yymm;

            ObjWorkSheet.Cells[7, 1] = report.CountMoCheck;
            ObjWorkSheet.Cells[7, 2] = report.CountMoCheckWithDefect;
            ObjWorkSheet.Cells[7, 3] = report.CountProporsals;
            ObjWorkSheet.Cells[7, 4] = report.CountProporsalsWithDefect;
            if (!string.IsNullOrEmpty(report.Notes))
                ObjWorkSheet.Cells[7, 6] = report.Notes;

            ObjWorkSheet.Cells[10, 2] = CurrentUser.UserName;
            ObjWorkSheet.Cells[12, 2] = CurrentUser.Director;


        }
    }
}
