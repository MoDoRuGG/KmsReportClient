using KmsReportClient.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using KmsReportClient.External;
using KmsReportClient.Support;

namespace KmsReportClient.Excel.Creator
{
    public class DynamicReportExcelCreator
    {
        protected Application ObjExcel;
        protected Workbook ObjWorkBook;
        protected Worksheet ObjWorkSheet;
        List<DynamicDataDto> _data;

        private string Filename;
        private readonly DynamicReport Report;

        public DynamicReportExcelCreator(string fileName, DynamicReport report)
        {
            this.Filename = fileName;
            this.Report = report;
        }

        public DynamicReportExcelCreator(string fileName, DynamicReport report,List<DynamicDataDto> data)
        {
            this.Filename = fileName;
            this.Report = report;
            _data = data;
        }


        public void FillReport()
        {
            int startRow = 5;
            int i = 1;
            int p = 1;
            bool IsGroup;
            bool checkRow;

            foreach (var page in Report.Page)
            {
                i = 1;
                startRow = 4;
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[p];
                ObjWorkSheet.Name = page.Key.Name;

                IsGroup = page.Value.Columns.Any(x => x.IsGroup);
                checkRow = page.Value.Rows.Any();


                if (checkRow)
                {
                    //Range group = (Range)ObjWorkSheet.Range[ObjWorkSheet.Cells[startRow, i], ObjWorkSheet.Cells[startRow + 1, i]].Cells;
                    //group.Merge(Type.Missing);
                    //ObjWorkSheet.Cells[startRow, i++] = "Наименование показателя";
                    //group.Borders.LineStyle = XlLineStyle.xlContinuous;


                    //Range group1 = (Range)ObjWorkSheet.Range[ObjWorkSheet.Cells[startRow, i], ObjWorkSheet.Cells[startRow + 1, i]].Cells;
                    //group1.Merge(Type.Missing);
                    //ObjWorkSheet.Cells[startRow, i++] = "№";
                    //group.Borders.LineStyle = XlLineStyle.xlContinuous;


                }


                foreach (var column in page.Value.Columns)
                {

                    if (!column.IsGroup)
                    {

                        Range group = (Range)ObjWorkSheet.Range[ObjWorkSheet.Cells[startRow, i], ObjWorkSheet.Cells[startRow + 1, i]].Cells;
                        group.Merge(Type.Missing);
                        ObjWorkSheet.Cells[startRow, i++] = column.Name;
                        group.Borders.LineStyle = XlLineStyle.xlContinuous;

                    }
                    else
                    {
                        Range group = (Range)ObjWorkSheet.Range[ObjWorkSheet.Cells[startRow, i], ObjWorkSheet.Cells[startRow, i + column.Columns.Count - 1]].Cells;
                        group.Merge(Type.Missing);
                        ObjWorkSheet.Cells[startRow, i] = column.Name;
                        group.Borders.LineStyle = XlLineStyle.xlContinuous;

                        foreach (var colInGroup in column.Columns)
                        {
                            ObjWorkSheet.Cells[startRow + 1, i++] = colInGroup.Name;
                        }

                    }

                    ((Range)ObjWorkSheet.Range[ObjWorkSheet.Cells[startRow + 1, 1], ObjWorkSheet.Cells[startRow + 1, i - 1]]).Borders.LineStyle = XlLineStyle.xlContinuous;
                }

                var rangeColumnsName = ((Range)ObjWorkSheet.Range[ObjWorkSheet.Cells[startRow, 1], ObjWorkSheet.Cells[startRow + 1, i - 1]]);
                rangeColumnsName.Interior.Color = XlRgbColor.rgbLightGray;
                rangeColumnsName.Font.Bold = true;

                startRow += 2;

                ((Range)ObjWorkSheet.Range[ObjWorkSheet.Cells[startRow, 1], ObjWorkSheet.Cells[startRow, i - 1]]).Borders.LineStyle = XlLineStyle.xlContinuous;
                foreach (var row in page.Value.Rows)
                {
                    ((Range)ObjWorkSheet.Range[ObjWorkSheet.Cells[startRow, 1], ObjWorkSheet.Cells[startRow, i - 1]]).Borders.LineStyle = XlLineStyle.xlContinuous;
                    ObjWorkSheet.Cells[startRow, 1] = row.Name;
                    ObjWorkSheet.Cells[startRow++, 2] = row.Index;

                }

                if (_data != null)
                {
                    var data = _data.Where(x => PositionSupport.GetPage(x.Position) == p - 1).ToList();
                    SetData(ObjWorkSheet, data, startRow, i - 1,checkRow);
                    //Console.WriteLine(startRow);

                }


                ObjWorkSheet.Columns.EntireColumn.AutoFit();
                SetStyle(ObjWorkSheet, i, IsGroup);

                if (p != Report.Page.Count)
                {
                    ObjWorkBook.Sheets.Add(Type.Missing, ObjWorkSheet, 1, Type.Missing);
                    p++;
                }
                


            }

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            ObjWorkSheet.Activate();
          



            // SetInfo(ObjWorkSheet,Report.NameReport,"");

        }

        public static void SetData(Worksheet worksheet, List<DynamicDataDto> data, int startRow, int countClmn,bool checkRow)
        {        
            foreach (var item in data)
            {
                int clmn = PositionSupport.GetColumn(item.Position) + 1;
                int row = PositionSupport.GetRow(item.Position) + 1;
                worksheet.Cells[row+5, clmn] = item.Value;

            }

        }



        public static void SetStyle(Worksheet worksheet, int count, bool isGroup)
        {
            for (int i = 1; i <= count; i++)
            {
                //worksheet.Columns[i].ColumnWidth += 8;
                worksheet.Cells[4, i].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                worksheet.Cells[5, i].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                worksheet.Cells[5, i].Style.VerticalAlignment = XlVAlign.xlVAlignCenter;
            }

            if (isGroup)
            {
                worksheet.Rows[4].RowHeight = 30;
            }

            worksheet.Rows[5].RowHeight = 21;

        }

        public void SetInfo(Worksheet worksheet, string reportName, string fillialName)
        {
            worksheet.Cells[1, 1] = "КАПИТАЛ МС";
            worksheet.Cells[2, 1] = String.Format($"Название отчёта:");
            worksheet.Cells[2, 2] = reportName;

        }


        public void CreateReport()
        {
            ObjExcel = new Application { DisplayAlerts = false };
            ObjWorkBook = ObjExcel.Workbooks.Add();
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];

            try
            {
                FillReport();
                ObjWorkBook.SaveAs(Filename, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);


            }
            finally
            {
                ObjExcel.Quit();
                ObjWorkBook = null;
                ObjWorkSheet = null;
                ObjExcel = null;
                GC.Collect();
            }
        }

    }
}
