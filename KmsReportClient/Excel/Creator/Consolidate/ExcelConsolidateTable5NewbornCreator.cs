using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using KmsReportClient.Support;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    public class ExcelConsolidateTable5NewbornCreator : ExcelBaseCreator<ConsolidateTable5Newborn[]>
    {
        private string _yymm;

        public ExcelConsolidateTable5NewbornCreator(
                                          string filename,
                                          string header,
                                          string filialName, string yymm) : base(filename, ExcelForm.consT5Newborn, header, filialName, false)
        {
            _yymm = yymm;
        }

        protected override void FillReport(ConsolidateTable5Newborn[] report, ConsolidateTable5Newborn[] yearReport)
        {
            int countReport = report.Length;
            int currentIndex = 7;

            string shortMonth = _yymm[2] == '0'
    ? _yymm.Substring(3)
    : _yymm.Substring(2);

            ObjWorkSheet.Cells[2, 2] = "Страхование новорожденных за " + shortMonth + " месяцев 20" + _yymm.Substring(0, 2) + " года";
            ObjWorkSheet.Cells[4, 4] = YymmUtils.ConvertYymmToDate(_yymm);

            foreach (var data in report)
            {
                ObjWorkSheet.Cells[currentIndex, 2] = data.RegionName;
                ObjWorkSheet.Cells[currentIndex, 3] = data.MarketShare;
                ObjWorkSheet.Cells[currentIndex, 4] = data.CountNewborn;
                ObjWorkSheet.Cells[currentIndex, 5] = data.CountMaterinityBills;
                ObjWorkSheet.Cells[currentIndex, 6] = data.ShareFromRegister;
                ObjWorkSheet.Cells[currentIndex, 7] = data.DeviationFromRegister;

                currentIndex++;
            }
        }
    }
}
