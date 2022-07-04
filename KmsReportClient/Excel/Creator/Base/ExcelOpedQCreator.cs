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
    public class ExcelOpedQCreator : ExcelBaseCreator<ReportOped>
    {
        

        List<string>_notPrintRow =  new List<string>{ "1", "1.3", "2", "2.3", "3", "3.3" };


        public ExcelOpedQCreator(
         string filename,
         ExcelForm reportName,
         string header,
         string filialName) : base(filename, reportName, header, filialName, false) { }


        protected override void FillReport(ReportOped report, ReportOped yearReport)
        {


            string year= 20 + report.Yymm.Substring(0, 2);          
            int q = Convert.ToInt32(report.Yymm.Last().ToString()) / 3;

            


            string HeaderText = $"Отчет о выполнении за {q} квартал {year}г.";
            ObjWorkSheet.Cells[7, 1] = HeaderText;
            ObjWorkSheet.Cells[9, 1] = $"подразделение(филиал) ООО Капитал  МС {FilialName}";

            ObjWorkSheet.Cells[27, 3] = CurrentUser.UserName;
            ObjWorkSheet.Cells[29, 3] = CurrentUser.Director;

            for (int i = 14; i <= 25; i++)
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
                }
               
            }
            
        }
    }
}
