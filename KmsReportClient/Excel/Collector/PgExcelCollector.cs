using System;
using System.Collections.Generic;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Support;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Collector
{
    class PgExcelCollector : ExcelBaseCollector
    {
        //private readonly string[] _columnsTable1 = { "2", "7" };
        private readonly string[] _columnsTable1 = { "2", "8", "9" };
        private readonly string[] _columnsTable2 = { "2", "5", "7", "8", "9", "10", "11" };
        private readonly string[] _columnsTable3 = { "2", "5", "7", "8", "9", "10", "11" };
        private readonly string[] _columnsTable4 = { "2", "5" };
        private readonly string[] _columnsTable5 = { "2", "4", "5", "6", "7", "8", "9" };
        private readonly string[] _columnsTable6 = { "2", "4", "5", "6", "7", "8", "9", "11", "12", "13", "14", "15", "16" };
        private readonly string[] _columnsTable8 = { "2", "4", "5", "6", "7", "8", "9", "11", "12", "13", "14", "15", "16" };
        private readonly string[] _columnsTable10 = { "2", "4" };
        private readonly string[] _columnsTable12 = { "2", "3", "4" };
        private readonly string[] _columnsTable13 = { "2", "4" };

        protected override void FillReport(string form, AbstractReport destReport, AbstractReport srcReport)
        {
            var destData = (destReport as ReportPg)?.ReportDataList.Single(r => r.Theme == form) ?? 
                           throw new Exception($"Can't find destReportDataList for form = {form}");
            var srcData = (srcReport as ReportPg)?.ReportDataList.Single(r => r.Theme == form) ?? 
                          throw new Exception($"Can't find srcReportDataList for form = {form}");
            destData.Data = srcData.Data;
        }

        protected override AbstractReport CollectReportData(string form)
        {
            var themeData = form switch {
                "Таблица 12" => FillTable12(),
                "Таблица 1" => FillTable1(form),
                "Таблица 2" => FillTable2(form),
                "Таблица 3" => FillTable2(form),
                "Таблица 4" => FillTable4(form),
                "Таблица 10" => FillTable10_13(form),
                "Таблица 13" => FillTable10_13(form),
                "Таблица 6" => FillTable68(form),
                "Таблица 8" => FillTable68(form),
                _ => FillTable5()
            };
            var report = new ReportPg { ReportDataList = new ReportPgDto[1] };
            report.ReportDataList[0] = new ReportPgDto
            {
                Theme = form,
                Data = themeData
            };
            return report;
        }

        private ReportPgDataDto[] FillTable4(string form)
        {
            var list = new List<ReportPgDataDto>();
            int countWorkSheet = ObjWorkBook.Worksheets.Count;

            for (int currentList = 1; currentList <= countWorkSheet; currentList++)
            {
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[currentList];
                int lastRow = GetLastRow();

                Dictionary<string, int> dictionary;
                int startRow = GetStartRow();

                switch (form)
                {
                    case "Таблица 4":
                        //startRow = 15;
                        dictionary = FindColumnIndexies(_columnsTable4, startRow - 1);
                        break;
                    case "Таблица 10":
                        //startRow = currentList == 1 ? 15 : 4;
                        dictionary = FindColumnIndexies(_columnsTable10, startRow - 1);
                        break;
                    default:
                        //startRow = 15;
                        dictionary = FindColumnIndexies(_columnsTable13, startRow - 1);
                        break;
                }

                for (int i = startRow; i <= lastRow; i++)
                {
                    var data = new ReportPgDataDto
                    {
                        Code = ObjWorkSheet.Cells[i, dictionary["2"]].Text,
                        CountSmo = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["5"]].Text)
                    };
                    list.Add(data);
                }
            }

            return list.ToArray();
        }

        private ReportPgDataDto[] FillTable12()
        {
            var list = new List<ReportPgDataDto>();

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            int lastRow = GetLastRow();
            int startRow = 16;
            var dictionary = FindColumnIndexies(_columnsTable12, startRow - 1);

            for (int i = startRow; i <= lastRow; i++)
            {
                var data = new ReportPgDataDto
                {
                    Code = ObjWorkSheet.Cells[i, dictionary["2"]].Text,
                    CountSmo = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["3"]].Text),
                    CountSmoAnother = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["4"]].Text)
                };
                list.Add(data);
            }

            return list.ToArray();
        }

        private ReportPgDataDto[] FillTable5()
        {
            var list = new List<ReportPgDataDto>();
            int countWorkSheet = ObjWorkBook.Worksheets.Count;

            for (int currentList = 1; currentList <= countWorkSheet; currentList++)
            {
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[currentList];
                int lastRow = GetLastRow();
                int startRow = currentList == 1 ? 16 : 5;
                var dictionary = FindColumnIndexies(_columnsTable5, startRow - 1);

                for (int i = startRow; i <= lastRow; i++)
                {
                    var data = new ReportPgDataDto
                    {
                        Code = ObjWorkSheet.Cells[i, dictionary["2"]].Text,
                        CountOutOfSmo = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["4"]].Text),
                        CountAmbulatory = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["5"]].Text),
                        CountDs = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["6"]].Text),
                        CountDsVmp = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["7"]].Text),
                        CountStac = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["8"]].Text),
                        CountStacVmp = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["9"]].Text)
                    };
                    list.Add(data);
                }
            }

            return list.ToArray();
        }

        private ReportPgDataDto[] FillTable2(string form)
        {
            var list = new List<ReportPgDataDto>();
            int countWorkSheet = ObjWorkBook.Worksheets.Count;

            for (int currentList = 1; currentList <= countWorkSheet; currentList++)
            {
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[currentList];
                int lastRow = GetLastRow();
                int startRow = currentList == 1 ? 16 : 5;

                Dictionary<string, int> dictionary = form == "Таблица 2" ?
                    FindColumnIndexies(_columnsTable2, startRow - 1) :
                    FindColumnIndexies(_columnsTable3, startRow - 1);

                for (int i = startRow; i <= lastRow; i++)
                {
                    var data = new ReportPgDataDto
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

        private ReportPgDataDto[] FillTable10_13(string form)
        {
            var list = new List<ReportPgDataDto>();
            int countWorkSheet = ObjWorkBook.Worksheets.Count;

            for (int currentList = 1; currentList <= countWorkSheet; currentList++)
            {
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[currentList];
                int lastRow = GetLastRow();

                Dictionary<string, int> dictionary;
                int startRow;
                string rowDataIndex;

                switch (form)
                {
                    case "Таблица 1":
                        if (currentList == 1)
                        {
                            continue;
                        }

                        startRow = currentList == 2 ? 8 : 6;
                        dictionary = FindColumnIndexies(_columnsTable1, startRow - 1);
                        rowDataIndex = "7";
                        break;
                    case "Таблица 10":
                        startRow = currentList == 1 ? 15 : 4;
                        dictionary = FindColumnIndexies(_columnsTable10, startRow - 1);
                        rowDataIndex = "4";
                        break;
                    default:
                        startRow = 15;
                        dictionary = FindColumnIndexies(_columnsTable13, startRow - 1);
                        rowDataIndex = "4";
                        break;
                }

                for (int i = startRow; i <= lastRow; i++)
                {
                    var data = new ReportPgDataDto
                    {
                        Code = ObjWorkSheet.Cells[i, dictionary["2"]].Text,
                        CountSmo = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary[rowDataIndex]].Text)
                    };
                    list.Add(data);
                }
            }

            return list.ToArray();
        }

        private ReportPgDataDto[] FillTable1(string form)
        {
            var list = new List<ReportPgDataDto>();
            int countWorkSheet = ObjWorkBook.Worksheets.Count;
            int startList = form == "Таблица 1" ? 2 : 1;

            for (int currentList = startList; currentList <= countWorkSheet; currentList++)
            {
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[currentList];

                if (!ObjWorkSheet.Name.ToLower().Contains("Pag".ToLower()))
                    continue;
                int lastRow = GetLastRow();

                Dictionary<string, int> dictionary;
                int startRow;
                string rowDataIndexFirst;
                string rowDataIndexSecond;
                startRow = GetStartRow();
               
                switch (form)
                {
                    case "Таблица 1":
                        //startRow = currentList == 2 ? 8 : 6;
                        rowDataIndexFirst = "8";
                        rowDataIndexSecond = "9";
                        dictionary = FindColumnIndexies(_columnsTable1, startRow - 1);
                        break;
                    default:
                        //startRow = 16;
                        rowDataIndexFirst = "5";
                        rowDataIndexSecond = "6";
                        dictionary = FindColumnIndexies(_columnsTable12, startRow - 1);
                        break;
                }

               
                for (int i = startRow; i <= lastRow; i++)
                {
                    var data = new ReportPgDataDto
                    {
                        Code = ObjWorkSheet.Cells[i, dictionary["2"]].Text,
                        CountSmo = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary[rowDataIndexFirst]].Text),
                        CountSmoAnother = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary[rowDataIndexSecond]].Text)
                       
                    };
                    list.Add(data);

                   
                }
            }

            return list.ToArray();
        }



        private ReportPgDataDto[] FillTable68(string form)
        {
            var list = new List<ReportPgDataDto>();
            int countWorkSheet = ObjWorkBook.Worksheets.Count;
            var columns = form == "Таблица 6" ? _columnsTable6 : _columnsTable8;

            for (int currentList = 1; currentList <= countWorkSheet; currentList++)
            {
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[currentList];
                int lastRow = GetLastRow();
                int startRow = currentList == 1 ? 16 : 5;
                var dictionary = FindColumnIndexies(columns, startRow - 1);

                for (int i = startRow; i <= lastRow; i++)
                {
                    var data = new ReportPgDataDto
                    {
                        Code = ObjWorkSheet.Cells[i, dictionary["2"]].Text,
                        CountOutOfSmo = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["4"]].Text),
                        CountAmbulatory = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["5"]].Text),
                        CountDs = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["6"]].Text),
                        CountDsVmp = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["7"]].Text),
                        CountStac = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["8"]].Text),
                        CountStacVmp = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["9"]].Text),
                        CountOutOfSmoAnother = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["11"]].Text),
                        CountAmbulatoryAnother = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["12"]].Text),
                        CountDsAnother = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["13"]].Text),
                        CountDsVmpAnother = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["14"]].Text),
                        CountStacAnother = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["15"]].Text),
                        CountStacVmpAnother = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["16"]].Text),
                    };

                    list.Add(data);
                }
            }

            return list.ToArray();
        }
    }
}