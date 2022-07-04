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
    public class ExcelConsoldateCPNPQ2Creator : ExcelBaseCreator<ConsolidateCPNP_Q_2[]>
    {

        string _header;
        public ExcelConsoldateCPNPQ2Creator(
       string filename,
       string header,
       string filialName) : base(filename, ExcelForm.consCPNPQ2, header, filialName, false)
        {
            _header = header;
        }


        protected override void FillReport(ConsolidateCPNP_Q_2[] report, ConsolidateCPNP_Q_2[] yearReport)
        {

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];

            int row = 9;
            CopyNullCells(ObjWorkSheet, report.Length, row);

            ObjWorkSheet.Cells[5, 3] = _header;

            int counter = 1;
            foreach(var data in report)
            {
                ObjWorkSheet.Cells[row, 1] = counter++;
                ObjWorkSheet.Cells[row, 2] = data.Filial;
                ObjWorkSheet.Cells[row, 3] = data.CountSporDoSuda;
                ObjWorkSheet.Cells[row, 4] = data.CountObosnZhalob;
                row++;

            }


        }
    }
}
