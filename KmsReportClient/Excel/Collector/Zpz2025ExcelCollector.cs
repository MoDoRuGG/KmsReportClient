using System;
using System.Collections.Generic;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Support;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Collector
{
    class Zpz2025ExcelCollector : ExcelBaseCollector
    {
        private readonly string[] _columnsTable1 = { "2", "8", "9","10" };
        private readonly string[] _columnsTable2 = { "2", "5", "7", "8", "9", "10", "11" };
        private readonly string[] _columnsTable3 = { "2", "5", "7", "8", "9", "10", "11" };
        private readonly string[] _columnsTable4 = { "2", "5" };
        private readonly string[] _columnsTable10 = { "2", "4" };

        protected override void FillReport(string form, AbstractReport destReport, AbstractReport srcReport)
        {
            var destData = (destReport as ReportZpz2025)?.ReportDataList.Single(r => r.Theme == form) ?? 
                           throw new Exception($"Can't find destReportDataList for form = {form}");
            var srcData = (srcReport as ReportZpz2025)?.ReportDataList.Single(r => r.Theme == form) ?? 
                          throw new Exception($"Can't find srcReportDataList for form = {form}");
            destData.Data = srcData.Data;
        }

        protected override AbstractReport CollectReportData(string form)
        {
            var themeData = form switch {
                "Таблица 1" => FillTable1(form),
                "Таблица 2" => FillTable2(form),
                "Таблица 3" => FillTable2(form),
                "Таблица 4" => FillTable4(form),
                "Таблица 10" => FillTable4(form),

            };
            var report = new ReportZpz2025 { ReportDataList = new ReportZpz2025Dto[1] };
            report.ReportDataList[0] = new ReportZpz2025Dto
            {
                Theme = form,
                Data = themeData
            };
            return report;
        }

        private ReportZpz2025DataDto[] FillTable4(string form)
        {
            var list = new List<ReportZpz2025DataDto>();
            int countWorkSheet = ObjWorkBook.Worksheets.Count;

            for (int currentList = 1; currentList <= countWorkSheet; currentList++)
            {
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[currentList];
                int lastRow = GetLastRow();

                Dictionary<string, int> dictionary;
                int startRow = GetStartRow();
                string k;
                switch (form)
                {
                    case "Таблица 4":
                        k = "5";    
                    //startRow = 15;
                        dictionary = FindColumnIndexies(_columnsTable4, startRow - 1);
                        break;
                    default:
                        //startRow = currentList == 1 ? 15 : 4;
                        k = "4";
                        dictionary = FindColumnIndexies(_columnsTable10, startRow - 1);
                        break;
                }

                for (int i = startRow; i <= lastRow; i++)
                {
                    var data = new ReportZpz2025DataDto
                    {
                        Code = ObjWorkSheet.Cells[i, dictionary["2"]].Text,
                        CountSmo = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary[k]].Text)
                    };
                    list.Add(data);
                }
            }

            return list.ToArray();
        }

        private ReportZpz2025DataDto[] FillTable2(string form)
        {
            var list = new List<ReportZpz2025DataDto>();
            int countWorkSheet = ObjWorkBook.Worksheets.Count;

            for (int currentList = 1; currentList <= countWorkSheet; currentList++)
            {
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[currentList];
                int lastRow = GetLastRow();
                int startRow = currentList == 1 ? 12 : 5;

                Dictionary<string, int> dictionary = form == "Таблица 2" ?
                    FindColumnIndexies(_columnsTable2, startRow - 1) :
                    FindColumnIndexies(_columnsTable3, startRow - 1);

                for (int i = startRow; i <= lastRow; i++)
                {
                    var data = new ReportZpz2025DataDto
                    {
                        Code = ObjWorkSheet.Cells[i, dictionary["2"]].Text,
                        CountSmo = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["5"]].Text),
                        CountInsured = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["7"]].Text),
                        CountInsuredRepresentative = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["8"]].Text),
                        CountTfoms = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["9"]].Text),
                        CountSmoAnother = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["10"]].Text),
                        CountProsecutor = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["11"]].Text)
                    };
                    list.Add(data);
                }
            }

            return list.ToArray();
        }

        private ReportZpz2025DataDto[] FillTable1(string form)
        {
            var list = new List<ReportZpz2025DataDto>();
            int countWorkSheet = ObjWorkBook.Worksheets.Count;

            for (int currentList = 1; currentList <= countWorkSheet; currentList++)
            {
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[currentList];
                int lastRow = GetLastRow();

                Dictionary<string, int> dictionary;
                int startRow;
                string rowDataIndex;

                if (form == "Таблица 1")
                {
                    if (currentList == 1)
                    {
                        continue;
                    }

                    startRow = currentList == 2 ? 8 : 6;
                    dictionary = FindColumnIndexies(_columnsTable1, startRow - 1);


                    for (int i = startRow; i <= lastRow; i++)
                    {
                        var data = new ReportZpz2025DataDto
                        {
                            Code = ObjWorkSheet.Cells[i, dictionary["2"]].Text,
                            CountSmo = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["8"]].Text),
                            CountSmoAnother = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["9"]].Text),
                            CountAssignment = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["10"]].Text)
                        };
                        list.Add(data);
                    }
                }
            }

            return list.ToArray();
        }
    }
}