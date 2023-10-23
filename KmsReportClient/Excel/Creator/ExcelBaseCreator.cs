using System;
using KmsReportClient.Global;
using KmsReportClient.Model.Enums;
using KmsReportClient.Support;
using Microsoft.Office.Interop.Excel;
using NLog;

namespace KmsReportClient.Excel.Creator
{
    public abstract class ExcelBaseCreator<T>
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();
        
        protected readonly ExcelForm ReportName;
        protected readonly string Filename;
        protected readonly string FilialName;
        protected readonly string Header;
        protected readonly bool IsNeedPdf;

        protected Application ObjExcel;
        protected Workbook ObjWorkBook;
        protected Worksheet ObjWorkSheet;

        public ExcelBaseCreator(string filename, ExcelForm reportName, string header, 
            string filialName, bool isNeedPdf)
        {
            this.Filename = filename;
            this.ReportName = reportName;
            this.Header = header;
            this.IsNeedPdf = isNeedPdf;
            this.FilialName = filialName;
 
            
        }

        public void CreateReport(T report, T yearReport)
        {
            ObjExcel = new Application { DisplayAlerts = false, 
                                       //Visible = true 
                                        };
            ObjWorkBook = ObjExcel.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + ReportName.GetDescription());
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];

            try
            {
                FillReport(report, yearReport);

                ObjWorkBook.SaveAs(Filename, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                if (IsNeedPdf)
                {
                    try
                    {
                        ObjWorkBook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, Filename.Replace("xlsx", "pdf"));
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex, "Error saving xlsx to pdf");
                    }
                }
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

        protected abstract void FillReport(T report, T yearReport);

        protected string GetPhoneCode(string phone)
        {
            if (!string.IsNullOrEmpty(phone))
            {
                int index = phone.IndexOf(")");
                if (index > 0)
                {
                    return phone.Substring(1, CurrentUser.DirectorPhone.IndexOf(")") - 1);
                }
                return phone.Substring(0, 3);
            }
            return "";
        }

        protected string GetPhoneNumber(string phone)
        {
            if (!string.IsNullOrEmpty(phone))
            {
                int index = phone.IndexOf(")");
                if (index > 0)
                {
                    return phone.Substring(CurrentUser.DirectorPhone.IndexOf(")") + 1);
                }
                return phone.Substring(3);
            }
            return "";
        }

        protected void CopyNullCells(Worksheet sheet, int count, int position)
        {
            for (int k = 1; k <= count - 2; k++)
            {
                var r = sheet.Range[position + ":" + position, Type.Missing];
                r.Copy(Type.Missing);
                r = sheet.Range[Convert.ToString(k + position) + ":" + Convert.ToString(k + position), Type.Missing];
                r.Insert(XlInsertShiftDirection.xlShiftDown);
            }
        }

        protected void CopyNullCells4Rows(Worksheet sheet, int count, int position)
        {
            for (int k = 1; k <= count - 2; k++)
            {
                var r = sheet.Range[position + ":" + (position+3), Type.Missing];
                r.Copy(Type.Missing);
                r = sheet.Range[Convert.ToString( position + 4) + ":" + Convert.ToString(position + 7), Type.Missing];
                r.Insert(XlInsertShiftDirection.xlShiftDown);
            }
        }


        protected void CopyNullCellsOped(Worksheet sheet, int count, int position)
        {
            int cntS = 10;
            int cntE = 12;
            for (int k = 1; k <= count - 1; k++)
            {

                var row = sheet.Range["A" + cntS + ":F" + cntE, Type.Missing];
                row.Copy(Type.Missing);
                cntS += 4;
                cntE += 4;
                row = sheet.Range["A" + cntS + ":F" + cntE, Type.Missing];
                row.Insert(XlInsertShiftDirection.xlShiftDown);

            }
        }

        protected void CopyNullCellsOpedU(Worksheet sheet, int count, int position)
        {
            int cntS = 10;
            int cntE = 12;
            for (int k = 1; k <= count - 1; k++)
            {

                var row = sheet.Range["A" + cntS + ":F" + cntE, Type.Missing];
                row.Copy(Type.Missing);
                cntS += 4;
                cntE += 4;
                row = sheet.Range["A" + cntS + ":F" + cntE, Type.Missing];
                row.Insert(XlInsertShiftDirection.xlShiftDown);

            }
        }
    }
}
