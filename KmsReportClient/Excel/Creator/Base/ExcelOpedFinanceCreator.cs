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
    public class ExcelOpedFinanceCreator : ExcelBaseCreator<ReportOpedFinance>
    {

        public ExcelOpedFinanceCreator(
         string filename,
         ExcelForm reportName,
         string header,
         string filialName) : base(filename, reportName, header, filialName, false)
        {
            
        }

        protected override void FillReport(ReportOpedFinance report, ReportOpedFinance yearReport)
        {


            string reportMonths = YymmUtils.GetMonth(report.Yymm.Substring(2, 2));
            string reportYear = 20 + report.Yymm.Substring(0, 2);

            //ObjWorkSheet.Cells[4, 2] = String.Format($@"{FilialName}");
            //ObjWorkSheet.Cells[5, 2] = String.Format($"Отчет о выполнении плановых значений ОПЭД за {reportMonths} {reportYear}");
            ObjWorkSheet.Cells[4, 1] = String.Format($@"{FilialName}");
            ObjWorkSheet.Cells[5, 5] = String.Format($" {reportMonths} {reportYear}");

            //ObjWorkSheet.Cells[20, 4] = CurrentUser.UserName;
            //ObjWorkSheet.Cells[25, 4] = CurrentUser.Director;
            ObjWorkSheet.Cells[11, 8] = CurrentUser.UserName;
            ObjWorkSheet.Cells[14, 8] = CurrentUser.Director;
            //ObjWorkSheet.Cells[28, 5] =  DateTime.Now.ToShortDateString();
            ObjWorkSheet.Cells[17, 8] = DateTime.Now.ToShortDateString();





            ////Заполнение статисческие данных

            ////Начинаем с 9 строки, т.к 1 строка считается формулой
            //int dataColumnNumber = 4;
            //for(int i = 8; i <= 19; i++)
            //{
            //    string excelRowNumber = ObjWorkSheet.Cells[i, 3].Value.ToString();
            //    //if (!String.IsNullOrEmpty(excelRowNumber))
            //    {
            //        var reportRow = report.ReportDataList.FirstOrDefault(x=>x.RowNum == excelRowNumber);
            //        if(reportRow != null  )
            //        {
            //            if (excelRowNumber != "1.")
            //            {
            //                ObjWorkSheet.Cells[i, dataColumnNumber] = reportRow.ValueFact;
            //                ObjWorkSheet.Cells[i, dataColumnNumber + 1] = reportRow.Notes;
            //            }
            //            else
            //            {
            //                ObjWorkSheet.Cells[i, dataColumnNumber + 1] = reportRow.Notes;

            //            }



            //        } 
            //    }
            //}



            int dataColumnNumber = 6;
            for (int i = 8; i <= 10; i++)
            {
                string excelRowNumber = ObjWorkSheet.Cells[i, 1].Value.ToString();
                //if (!String.IsNullOrEmpty(excelRowNumber))
                {
                    var reportRow = report.ReportDataList.FirstOrDefault(x => x.RowNum == excelRowNumber);
                    if (reportRow != null)
                    {
                            ObjWorkSheet.Cells[i, dataColumnNumber] = reportRow.ValueFact;
                            ObjWorkSheet.Cells[i, dataColumnNumber + 1] = reportRow.Notes;
                    }
                }
            }
        }
    }
}
