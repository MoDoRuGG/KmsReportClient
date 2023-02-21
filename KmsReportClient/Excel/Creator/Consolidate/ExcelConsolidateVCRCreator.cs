using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    public class ExcelConsolidateVCRCreator : ExcelBaseCreator<ConsolidateVCR[]>
    {
        private string _yymm;

        public ExcelConsolidateVCRCreator(
                                          string filename,
                                          string header,
                                          string filialName, string yymm) : base(filename, ExcelForm.consVCR, header, filialName, false)
        {
            _yymm = yymm;
        }

        protected override void FillReport(ConsolidateVCR[] report, ConsolidateVCR[] yearReport)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];

            for (int i = 13; i <= 36; i++)
            {
                string rowNum = Convert.ToString(ObjWorkSheet.Cells[i, 1].Value);
                var data = report.FirstOrDefault(x=> x.RowNum == rowNum);
                if(data != null)
                {
                    ObjWorkSheet.Cells[i, 3] = data.ExpertWithEducation;
                    ObjWorkSheet.Cells[i, 4] = data.ExpertWithoutEducation;
                    ObjWorkSheet.Cells[i, 5] = data.Total;
                }

            }


        }
    }
}
