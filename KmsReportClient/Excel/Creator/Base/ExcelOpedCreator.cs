using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KmsReportClient.External;
using KmsReportClient.Model;
using KmsReportClient.Model.Enums;
using KmsReportClient.Support;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Base
{
    public class ExcelOpedCreator : ExcelBaseCreator<ReportOped>
    {
        private Dictionary<string, List<CellModel>> staticNormativ;


        List<CellModel> beforeJunyNormativ = new List<CellModel>()
        {
            new CellModel
            {
                Row = 17,
                Column =3,
                Value = 0.8M,

            },
            new CellModel
            {
                Row = 17,
                Column =4,
                Value = 8M,

            },

            new CellModel
            {
                Row = 17,
                Column = 5,
                Value = 8M,

            },
               new CellModel
               {
                   Row = 17,
                   Column = 6,
                   Value = 3M,

               },


                new CellModel
                {
                    Row = 18,
                    Column = 3,
                    Value = 0.5M,

                },
            new CellModel
            {
                Row = 18,
                Column = 4,
                Value = 5M,

            },

            new CellModel
            {
                Row = 18,
                Column = 5,
                Value = 3M,

            },
             new CellModel
             {
                 Row = 18,
                 Column = 6,
                 Value = 1.5M,

             },


        };

        List<CellModel> AfterJunyNormativ = new List<CellModel>()
        {
            new CellModel
            {
                Row = 17,
                Column =3,
                Value = 0.5M,

            },
            new CellModel
            {
                Row = 17,
                Column =4,
                Value = 6M,

            },

            new CellModel
            {
                Row = 17,
                Column = 5,
                Value = 6M,

            },
               new CellModel
               {
                   Row = 17,
                   Column = 6,
                   Value = 2M,

               },


                new CellModel
                {
                    Row = 18,
                    Column = 3,
                    Value = 0.2M,

                },
            new CellModel
            {
                Row = 18,
                Column = 4,
                Value = 3M,

            },

            new CellModel
            {
                Row = 18,
                Column = 5,
                Value = 1.5M,

            },
             new CellModel
             {
                 Row = 18,
                 Column = 6,
                 Value = 0.5M,

             },


        };



        public ExcelOpedCreator(
          string filename,
          ExcelForm reportName,
          string header,
          string filialName) : base(filename, reportName, header, filialName, false)
        {
            staticNormativ = new Dictionary<string, List<CellModel>>();
            staticNormativ.Add("2105", beforeJunyNormativ);
            staticNormativ.Add("2106", AfterJunyNormativ);
        }

        protected override void FillReport(ReportOped report, ReportOped yearReport)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            ObjWorkSheet.Name = Header;

            string reportMonths = YymmUtils.GetMonth(report.Yymm.Substring(2, 2));
            string reportYear = 20 + report.Yymm.Substring(0, 2);

            ObjWorkSheet.Cells[11, 2] = String.Format($"ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ 'КАПИТАЛ МЕДИЦИНСКОЕ СТРАХОВАНИЕ' ({FilialName})");
            ObjWorkSheet.Cells[9, 2] = String.Format($"Отчет о выполнении нормативов объемов экспертиз за {reportMonths} {reportYear} года");

            //Заполнение статисческие данных

            string yymmForNormativ = Convert.ToInt32(report.Yymm) <= 2105 ? "2105" : "2106" ;

         
            List<CellModel> normativ;
            if (staticNormativ.TryGetValue(yymmForNormativ, out normativ))
            {
                foreach(var n in normativ)
                {
                    ObjWorkSheet.Cells[n.Row, n.Column] = n.Value;
                }
            }
            

            for (int i = 16; i <= 24; i++)
            {
                string exRowNum = Convert.ToString(ObjWorkSheet.Cells[i, 1].Value);
                var rowData = report.ReportDataList.SingleOrDefault(x => x.RowNum == exRowNum);
                if (rowData != null)
                {
                    SetRow(rowData, i);
                }
            }


            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[2];
            ObjWorkSheet.Name = "Свод";

            reportMonths = YymmUtils.GetMonth(report.Yymm.Substring(2, 2));
            reportYear = 20 + report.Yymm.Substring(0, 2);

            ObjWorkSheet.Cells[11, 2] = String.Format($"ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ 'КАПИТАЛ МЕДИЦИНСКОЕ СТРАХОВАНИЕ' ({FilialName})");
            ObjWorkSheet.Cells[9, 2] = String.Format($"Отчет о выполнении нормативов объемов экспертиз с Января по {reportMonths} {reportYear} года");

            //Заполнение статисческие данных

            yymmForNormativ = Convert.ToInt32(report.Yymm) <= 2105 ? "2105" : "2106";


            if (staticNormativ.TryGetValue(yymmForNormativ, out normativ))
            {
                foreach (var n in normativ)
                {
                    ObjWorkSheet.Cells[n.Row, n.Column] = n.Value;
                }
            }


            for (int i = 16; i <= 24; i++)
            {
                string exRowNum = Convert.ToString(ObjWorkSheet.Cells[i, 1].Value);
                var rowData = yearReport.ReportDataList.SingleOrDefault(x => x.RowNum == exRowNum);
                if (rowData != null)
                {
                    SetRow(rowData, i);
                }
            }
        }


        private void SetRow(ReportOpedDto data, int rowNum)
        {
            var formulaRows = new int[] { 17, 18, 21, 22, 23, 24 };
            if (!formulaRows.Contains(rowNum))
            {
                ObjWorkSheet.Cells[rowNum, 3] = data.App;
                ObjWorkSheet.Cells[rowNum, 4] = data.Ks;
                ObjWorkSheet.Cells[rowNum, 5] = data.Ds;
                ObjWorkSheet.Cells[rowNum, 6] = data.Smp;
            }
            ObjWorkSheet.Cells[rowNum, 7] = data.Notes;







        }
    }
}
