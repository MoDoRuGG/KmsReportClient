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
    public class ExcelConsolidateOpedCreator : ExcelBaseCreator<ConsolidateOped[]>
    {

        struct FormulaOped
        {
            public char Source { get; set; }
            public char Formula { get; set; }

            public int ClmnFormula { get; set; }


        }


        private string _start;
        private string _end;
        public ExcelConsolidateOpedCreator(
       string filename,
       string header,
       string filialName, string start, string end) : base(filename, ExcelForm.consOped, header, filialName, false)
        {
            _start = start;
            _end = end;
        }

        protected override void FillReport(ConsolidateOped[] report, ConsolidateOped[] yearReport)
        {
            FillOped(report);
        }

        public void FillOped(ConsolidateOped[] reports)
        {
            var opeds = reports.Select(r => new { r.Filial, r.Mee, r.Ekmp }).ToList();

            int countReport = opeds.Count;
            int currentIndex = 10;

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            CopyNullCellsOped(ObjWorkSheet, countReport, 5);

            ObjWorkSheet.Cells[2, 3] = $"{_start} - {_end}";

            int counter = 1;
            foreach (var data in opeds)
            {
                ObjWorkSheet.Cells[currentIndex, 1] = counter++;
                ObjWorkSheet.Cells[currentIndex, 2] = data.Filial;

                //МЭЭ
                ObjWorkSheet.Cells[currentIndex + 1, 3] = data.Mee.app;
                ObjWorkSheet.Cells[currentIndex + 1, 4] = data.Mee.ks;
                ObjWorkSheet.Cells[currentIndex + 1, 5] = data.Mee.ds;
                ObjWorkSheet.Cells[currentIndex + 1, 6] = data.Mee.smp;

                //ЭКМП
                ObjWorkSheet.Cells[currentIndex + 2, 3] = data.Ekmp.app;
                ObjWorkSheet.Cells[currentIndex + 2, 4] = data.Ekmp.ks;
                ObjWorkSheet.Cells[currentIndex + 2, 5] = data.Ekmp.ds;
                ObjWorkSheet.Cells[currentIndex + 2, 6] = data.Ekmp.smp;

                currentIndex += 4;

            }

            List<FormulaOped> formulas = new List<FormulaOped>();
            formulas.Add(new FormulaOped
            {
                Source = 'C',
                Formula = 'J',
                ClmnFormula = 10

            });

            formulas.Add(new FormulaOped
            {
                Source = 'D',
                Formula = 'K',
                ClmnFormula = 11
            });

            formulas.Add(new FormulaOped
            {
                Source = 'E',
                Formula = 'L',
                ClmnFormula = 12
            });

            formulas.Add(new FormulaOped
            {
                Source = 'F',
                Formula = 'M',
                ClmnFormula = 13
            });

            foreach (var f in formulas)
            {
                int cnt = 11;
                string mee = "=(";
                string emkp = "=(";
                foreach (var op in opeds)
                {
                    mee += f.Source.ToString() + cnt + "+";
                    emkp += f.Source.ToString() + (cnt + 1) + "+";

                    cnt += 4;
                }

                mee = mee.Remove(mee.Length - 1) + ")" + "/" + opeds.Count;
                emkp = emkp.Remove(emkp.Length - 1) + ")" + "/" + opeds.Count;

                ObjWorkSheet.Cells[5, f.ClmnFormula].Formula = mee;
                ObjWorkSheet.Cells[6, f.ClmnFormula].Formula = emkp;
               
            }


        }

    }
}
