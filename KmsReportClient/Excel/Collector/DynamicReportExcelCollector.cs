using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KmsReportClient.External;
using KmsReportClient.Model;
using KmsReportClient.Report.Basic;
using Microsoft.Office.Interop.Excel;
using NLog;

namespace KmsReportClient.Excel.Collector
{
    public class DynamicReportExcelCollector
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        protected Application ObjExcel;
        protected Workbook ObjWorkBook;
        protected Worksheet ObjWorkSheet;

        public DynamicReportExcelCollector()
        {

        }

        public void Collect(string filename, DynamicReportProcessor processor, DynamicReport report)
        {
            ObjExcel = new Application();
            ObjWorkBook = ObjExcel.Workbooks.Open(filename);

            try
            {
                var list = CollectReportData(report);
                processor.data.Clear();
                processor.data = list;

            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка обработки входного документа");
                throw;
            }
            finally
            {
                ObjExcel.Quit();
                ObjWorkBook = null;
                ObjWorkSheet = null;
                ObjExcel = null;
            }
        }

        private List<DynamicDataDto> CollectReportData(DynamicReport report)
        {
            List<DynamicDataDto> data = new List<DynamicDataDto>();
            int lastRow;
            int lastClmn;
            int pageIndex = 0;
            int clmnInc = 0;

            foreach (Worksheet displayWorksheet in ObjWorkBook.Worksheets)
            {
                var page = report.Page.ElementAt(pageIndex);

                if (page.Value.Rows.Any())
                {
                    clmnInc = 2;
                }
                else
                {
                    clmnInc = 0;
                }

                lastRow = GetLastRow(displayWorksheet);
                lastClmn = GetLastClmn(displayWorksheet);
                for (int row = 6; row <= lastRow; row++)
                {
                    for (int clmn = 1 + clmnInc; clmn <= lastClmn; clmn++)
                    {
                        Console.WriteLine($"page={pageIndex} row={row} clmn={clmn}");
                        data.Add(new DynamicDataDto
                        {
                            Position = GetPosition(pageIndex, clmn - 1, row - 6),
                            Value = displayWorksheet.Cells[row, clmn].Text

                        });
                    }
                }
                pageIndex++;
            }

            return data;
        }

        public string GetPosition(int page, int column, int row) => String.Format($"P{page}C{column}R{row}");

        protected int GetLastRow(Worksheet sheet) => sheet.Cells[sheet.Rows.Count, "A"].End[XlDirection.xlUp].Row;
      

        protected int GetLastClmn(Worksheet sheet) => sheet.Cells[6, sheet.Columns.Count].End[XlDirection.xlToLeft].Column;
    }
}
