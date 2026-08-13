using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Base
{
    public class Excel140nCreator : ExcelBaseCreator<Report140n>
    {
        private EndpointSoap _client;
        private string _regionCode;

        public Excel140nCreator(
          string filename,
          ExcelForm reportName,
          string header,
          string filialName, EndpointSoap client, string regionCode) : base(filename, reportName, header, filialName, false)
        {
            _client = client;
            _regionCode = regionCode;
        }

        protected override void FillReport(Report140n report, Report140n yearReport)
        {
            int sheet = 1;

            foreach (var theme in report.ReportDataList)
            {
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[sheet];

                if (theme.Data != null)
                {
                    ObjWorkSheet.Cells[7, 1] = FilialName;

                    // 1. Ведение персонифицированного учета
                    ObjWorkSheet.Cells[7, 2] = theme.Data.CZLdost;
                    ObjWorkSheet.Cells[7, 3] = theme.Data.CZLsmo;

                    // 2. Эффективность сопровождения
                    ObjWorkSheet.Cells[7, 5] = theme.Data.KSErez;
                    ObjWorkSheet.Cells[7, 6] = theme.Data.KSE;

                    // 3. Диспансеризация взрослого населения
                    ObjWorkSheet.Cells[7, 8] = theme.Data.PPMinadvn;
                    ObjWorkSheet.Cells[7, 9] = theme.Data.Iidvn;

                    // 4. Диспансерное наблюдение
                    ObjWorkSheet.Cells[7, 11] = theme.Data.PPMinfdn;
                    ObjWorkSheet.Cells[7, 12] = theme.Data.Iidn;

                    // 5. Защита прав (судебный и досудебный)
                    ObjWorkSheet.Cells[7, 14] = theme.Data.KOJdosud;
                    ObjWorkSheet.Cells[7, 15] = theme.Data.KOJsud;
                    ObjWorkSheet.Cells[7, 16] = theme.Data.KOJzl;

                    // 6. Эффективность защиты прав
                    ObjWorkSheet.Cells[7, 18] = theme.Data.KOJzlsmo;
                    ObjWorkSheet.Cells[7, 19] = theme.Data.CZLsmo; // дублируется по структуре отчета

                    // 7. Авансирование МО
                    ObjWorkSheet.Cells[7, 21] = theme.Data.KZAsobl;
                    ObjWorkSheet.Cells[7, 22] = theme.Data.KZAvsego;

                    // 8. Контроль использования средств (дебиторка)
                    ObjWorkSheet.Cells[7, 24] = theme.Data.DT;
                    ObjWorkSheet.Cells[7, 25] = theme.Data.Scpo;

                    // 9. Эффективность экспертной деятельности
                    ObjWorkSheet.Cells[7, 27] = theme.Data.KEKMPpodtv;
                    ObjWorkSheet.Cells[7, 28] = theme.Data.KEKMPtfoms;

                    // 10. Качество контроля СМО
                    ObjWorkSheet.Cells[7, 30] = theme.Data.KZSMOpodtv;
                    ObjWorkSheet.Cells[7, 31] = theme.Data.KPMOtfoms;

                    if (!string.IsNullOrEmpty(report.Yymm) && report.Yymm.EndsWith("03"))
                    {
                        ObjWorkSheet.Cells[7, 26] = "=(1-(X7/Y7)-1/12*190%)*100";
                    }
                    else
                    {
                        ObjWorkSheet.Cells[7, 26] = "=(1-(X7/Y7)-1/12)*100";
                    }
                }

                sheet++;
            }
        }
    }
}