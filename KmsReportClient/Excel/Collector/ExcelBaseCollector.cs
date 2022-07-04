using System;
using System.Collections.Generic;
using KmsReportClient.External;
using Microsoft.Office.Interop.Excel;
using NLog;

namespace KmsReportClient.Excel.Collector
{
    abstract class ExcelBaseCollector: IExcelCollector
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        protected Application ObjExcel;
        protected Workbook ObjWorkBook;
        protected Worksheet ObjWorkSheet;

        public void Collect(string filename, string form, AbstractReport report)
        {
            ObjExcel = new Application();
            ObjWorkBook = ObjExcel.Workbooks.Open(filename);
            //asdasd

            try
            {
                var list = CollectReportData(form);
                FillReport(form, report, list);
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

        protected int GetLastRow() =>
            ObjWorkSheet.Cells.Find(
                "*",
                ObjWorkSheet.Cells[1, 1],
                XlFindLookIn.xlFormulas,
                XlLookAt.xlPart,
                XlSearchOrder.xlByRows,
                XlSearchDirection.xlPrevious).Row;


        protected int GetStartRow()
        {
            int lastRow = GetLastRow();
            int startRow = 0;
            for(int i = 1;i<=lastRow; i++)
            {
                if (Convert.ToString(ObjWorkSheet.Cells[i,1].Text)=="1")
                {
                    startRow = i+1;
                    break;
                }
            }

            return startRow;

        }

        protected Dictionary<string, int> FindColumnIndexies(string[] columns, int rowNum)
        {
            var dictionary = new Dictionary<string, int>();
            int i = 1;
            foreach (var column in columns)
            {
                bool isFindIndex = false;
                while (!isFindIndex)
                {
                    string text = ObjWorkSheet.Cells[rowNum, i].Text;
                    if (text == column)
                    {
                        dictionary.Add(column, i);
                        isFindIndex = true;
                    }
                    i++;
                }
            }
            return dictionary;
        }

        protected abstract void FillReport(string form, AbstractReport destReport, AbstractReport srcReport);

        protected abstract AbstractReport CollectReportData(string form);
    }
}