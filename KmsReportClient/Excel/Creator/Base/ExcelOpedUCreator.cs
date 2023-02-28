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
    public class ExcelOpedUCreator : ExcelBaseCreator<ReportOpedU>
    {
        

        List<string>_notPrintRow =  new List<string>{"1.3","2.3","3.3"};


        public ExcelOpedUCreator(
         string filename,
         ExcelForm reportName,
         string header,
         string filialName) : base(filename, reportName, header, filialName, false) { }


        protected override void FillReport(ReportOpedU report, ReportOpedU yearReport)
        {
            string year = 20 + report.Yymm.Substring(0, 2);
            int q = Convert.ToInt32(report.Yymm.Substring(2, 2)) / 3;

            string HeaderText = $"Отчет о выполнении за {q} квартал {year}г.";
            ObjWorkSheet.Cells[3, 1] = HeaderText;
            ObjWorkSheet.Cells[5, 1] = $"подразделение(филиал) ООО Капитал  МС {FilialName}";

            for (int i = 11; i <= 19; i++)
            {
                var exRowNum = Convert.ToString(ObjWorkSheet.Cells[i, 1].Value);
                var rowData = report.ReportDataList.SingleOrDefault(x => x.RowNum == exRowNum);
                if (rowData != null && !_notPrintRow.Contains(exRowNum))
                {
                    ObjWorkSheet.Cells[i, 3] = rowData.App;
                    ObjWorkSheet.Cells[i, 4] = rowData.Ks;
                    ObjWorkSheet.Cells[i, 5] = rowData.Ds;
                    ObjWorkSheet.Cells[i, 6] = rowData.Smp;
                    ObjWorkSheet.Cells[i, 7] = rowData.Notes;
                    ObjWorkSheet.Cells[i, 8] = rowData.NotesGoodReason;
                }
               
            }
            
        }
    }
}
