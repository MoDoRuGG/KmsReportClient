using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelConsolidateOpedUnplannedCreator : ExcelBaseCreator<CReportOpedUnplanned[]>
    {
        private const int StartPosition = 11;
        private readonly List<KmsReportDictionary> _regions;
        public ExcelConsolidateOpedUnplannedCreator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.consOpedU, header, filialName, false) { }

        protected override void FillReport(CReportOpedUnplanned[] report, CReportOpedUnplanned[] yearReport)
        {
            int countReport = report.Length;
            int currentIndex = StartPosition;

            int ind = 0;

            while (ind < 252 && currentIndex < 420)
            {


                

                for (int i = currentIndex; i <= currentIndex + 9; i++)
                {
                    //.Filial = _regions.Single(j => j.Key == d.Filial).Value;
                    //string exRowNum = Convert.ToString(ObjWorkSheet.Cells[i, 2].Value);
                    var formulaRows = new int[] { currentIndex, currentIndex + 1, currentIndex + 3, currentIndex + 4, currentIndex + 6, currentIndex + 7 };
                    if (formulaRows.Contains(i))
                    {
                        ObjWorkSheet.Cells[i, 4] = report[ind].App;
                        ObjWorkSheet.Cells[i, 5] = report[ind].Ks;
                        ObjWorkSheet.Cells[i, 6] = report[ind].Ds;
                        ObjWorkSheet.Cells[i, 7] = report[ind].Smp;
                        ObjWorkSheet.Cells[i, 8] = report[ind].Notes;
                        ObjWorkSheet.Cells[i, 9] = report[ind].NotesGoodReason;
                        ind++;
                    }
                    
                   
                }
                currentIndex = currentIndex + 10;
            }
        }
    }
}

//private void SetRow(CReportOpedUnplanned data, int rowNum)
//{
//    var formulaRows = new int[] { currentIn, 18, 21, 22, 23, 24 };
//    if (!formulaRows.Contains(rowNum))
//    {
//        ObjWorkSheet.Cells[rowNum, 3] = data.App;
//        ObjWorkSheet.Cells[rowNum, 4] = data.Ks;
//        ObjWorkSheet.Cells[rowNum, 5] = data.Ds;
//        ObjWorkSheet.Cells[rowNum, 6] = data.Smp;
//    }
//    ObjWorkSheet.Cells[rowNum, 7] = data.Notes;
//}



