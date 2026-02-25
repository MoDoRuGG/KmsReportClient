using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    public class ExcelConsolidateZpzTable10FilialCreator : ExcelBaseCreator<ConsolidateZpzTable10Filial[]>
    {
        private string _yymm;

        public ExcelConsolidateZpzTable10FilialCreator(
                                          string filename,
                                          string header,
                                          string filialName, string yymm) : base(filename, ExcelForm.Zpz10ConsFilial, header, filialName, false)
        {
            _yymm = yymm;
        }

        protected override void FillReport(ConsolidateZpzTable10Filial[] report, ConsolidateZpzTable10Filial[] yearReport)
        {
            int countReport = report.Length;
            int currentIndex = 3;
            CopyNullCells(ObjWorkSheet, countReport+1, 3);

            foreach (var data in report)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = data.RegionName;
                ObjWorkSheet.Cells[currentIndex, 2] = data._1;
                ObjWorkSheet.Cells[currentIndex, 3] = data._11;
                ObjWorkSheet.Cells[currentIndex, 4] = data._12;
                ObjWorkSheet.Cells[currentIndex, 5] = data._13;
                ObjWorkSheet.Cells[currentIndex, 6] = data._14;
                ObjWorkSheet.Cells[currentIndex, 7] = data._2;
                ObjWorkSheet.Cells[currentIndex, 8] = data._21;
                ObjWorkSheet.Cells[currentIndex, 9] = data._22;
                ObjWorkSheet.Cells[currentIndex, 10] = data._23;
                ObjWorkSheet.Cells[currentIndex, 11] = data._24;
                ObjWorkSheet.Cells[currentIndex, 12] = data._3;
                ObjWorkSheet.Cells[currentIndex, 13] = data._31;
                ObjWorkSheet.Cells[currentIndex, 14] = data._32;
                ObjWorkSheet.Cells[currentIndex, 15] = data._33;
                ObjWorkSheet.Cells[currentIndex, 16] = data._34;
                ObjWorkSheet.Cells[currentIndex, 17] = data._4;
                ObjWorkSheet.Cells[currentIndex, 18] = data._41;
                ObjWorkSheet.Cells[currentIndex, 19] = data._411;
                ObjWorkSheet.Cells[currentIndex, 20] = data._412;
                ObjWorkSheet.Cells[currentIndex, 21] = data._413;
                ObjWorkSheet.Cells[currentIndex, 22] = data._414;
                ObjWorkSheet.Cells[currentIndex, 23] = data._42;
                ObjWorkSheet.Cells[currentIndex, 24] = data._421;
                ObjWorkSheet.Cells[currentIndex, 25] = data._422;
                ObjWorkSheet.Cells[currentIndex, 26] = data._423;
                ObjWorkSheet.Cells[currentIndex, 27] = data._424;
                ObjWorkSheet.Cells[currentIndex, 28] = data._43;
                ObjWorkSheet.Cells[currentIndex, 29] = data._431;
                ObjWorkSheet.Cells[currentIndex, 30] = data._432;
                ObjWorkSheet.Cells[currentIndex, 31] = data._433;
                ObjWorkSheet.Cells[currentIndex, 32] = data._434;
                ObjWorkSheet.Cells[currentIndex, 33] = data._44;
                ObjWorkSheet.Cells[currentIndex, 34] = data._441;
                ObjWorkSheet.Cells[currentIndex, 35] = data._442;
                ObjWorkSheet.Cells[currentIndex, 36] = data._443;
                ObjWorkSheet.Cells[currentIndex, 37] = data._444;
                ObjWorkSheet.Cells[currentIndex, 38] = data._45;
                ObjWorkSheet.Cells[currentIndex, 39] = data._451;
                ObjWorkSheet.Cells[currentIndex, 40] = data._452;
                ObjWorkSheet.Cells[currentIndex, 41] = data._453;
                ObjWorkSheet.Cells[currentIndex, 42] = data._454;
                ObjWorkSheet.Cells[currentIndex, 43] = data._46;
                ObjWorkSheet.Cells[currentIndex, 44] = data._461;
                ObjWorkSheet.Cells[currentIndex, 45] = data._462;
                ObjWorkSheet.Cells[currentIndex, 46] = data._463;
                ObjWorkSheet.Cells[currentIndex, 47] = data._464;
                ObjWorkSheet.Cells[currentIndex, 48] = data._5;
                ObjWorkSheet.Cells[currentIndex, 49] = data._51;
                ObjWorkSheet.Cells[currentIndex, 50] = data._511;
                ObjWorkSheet.Cells[currentIndex, 51] = data._512;
                ObjWorkSheet.Cells[currentIndex, 52] = data._513;
                ObjWorkSheet.Cells[currentIndex, 53] = data._514;
                ObjWorkSheet.Cells[currentIndex, 54] = data._52;
                ObjWorkSheet.Cells[currentIndex, 55] = data._521;
                ObjWorkSheet.Cells[currentIndex, 56] = data._522;
                ObjWorkSheet.Cells[currentIndex, 57] = data._523;
                ObjWorkSheet.Cells[currentIndex, 58] = data._524;
                ObjWorkSheet.Cells[currentIndex, 59] = data._53;
                ObjWorkSheet.Cells[currentIndex, 60] = data._531;
                ObjWorkSheet.Cells[currentIndex, 61] = data._532;
                ObjWorkSheet.Cells[currentIndex, 62] = data._533;
                ObjWorkSheet.Cells[currentIndex, 63] = data._534;
                ObjWorkSheet.Cells[currentIndex, 64] = data._54;
                ObjWorkSheet.Cells[currentIndex, 65] = data._541;
                ObjWorkSheet.Cells[currentIndex, 66] = data._542;
                ObjWorkSheet.Cells[currentIndex, 67] = data._543;
                ObjWorkSheet.Cells[currentIndex, 68] = data._544;
                ObjWorkSheet.Cells[currentIndex, 69] = data._55;
                ObjWorkSheet.Cells[currentIndex, 70] = data._551;
                ObjWorkSheet.Cells[currentIndex, 71] = data._552;
                ObjWorkSheet.Cells[currentIndex, 72] = data._553;
                ObjWorkSheet.Cells[currentIndex, 73] = data._554;
                ObjWorkSheet.Cells[currentIndex, 74] = data._56;
                ObjWorkSheet.Cells[currentIndex, 75] = data._561;
                ObjWorkSheet.Cells[currentIndex, 76] = data._562;
                ObjWorkSheet.Cells[currentIndex, 77] = data._563;
                ObjWorkSheet.Cells[currentIndex, 78] = data._564;
                ObjWorkSheet.Cells[currentIndex, 79] = data._6;
                ObjWorkSheet.Cells[currentIndex, 80] = data._61;
                ObjWorkSheet.Cells[currentIndex, 81] = data._62;
                ObjWorkSheet.Cells[currentIndex, 82] = data._63;
                ObjWorkSheet.Cells[currentIndex, 83] = data._64;
                ObjWorkSheet.Cells[currentIndex, 84] = data._65;
                ObjWorkSheet.Cells[currentIndex, 85] = data._66;
                ObjWorkSheet.Cells[currentIndex, 86] = data._67;
                ObjWorkSheet.Cells[currentIndex, 87] = data._7;
                ObjWorkSheet.Cells[currentIndex, 88] = data._71;
                ObjWorkSheet.Cells[currentIndex, 89] = data._72;
                ObjWorkSheet.Cells[currentIndex, 90] = data._73;
                ObjWorkSheet.Cells[currentIndex, 91] = data._74;
                ObjWorkSheet.Cells[currentIndex, 92] = "X";
                ObjWorkSheet.Cells[currentIndex, 93] = data._76;
                ObjWorkSheet.Cells[currentIndex, 94] = data._77;
                ObjWorkSheet.Cells[currentIndex, 95] = data._78;
                ObjWorkSheet.Cells[currentIndex, 96] = data._8;
                ObjWorkSheet.Cells[currentIndex, 97] = data._81;
                ObjWorkSheet.Cells[currentIndex, 98] = data._82;
                ObjWorkSheet.Cells[currentIndex, 99] = data._83;
                ObjWorkSheet.Cells[currentIndex, 100] = data._84;
                ObjWorkSheet.Cells[currentIndex, 101] = data._85;
                ObjWorkSheet.Cells[currentIndex, 102] = data._86;
                
                currentIndex++;
            }
        }
    }
}
