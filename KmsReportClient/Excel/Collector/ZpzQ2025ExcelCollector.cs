using System;
using System.Collections.Generic;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Support;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Collector
{
    class ZpzQ2025ExcelCollector : ExcelBaseCollector
    {
        private readonly string[] _columnsTable5A = { "2", "4", "5", "6", "7", "8", "9" };
        private readonly string[] _columnsTable6 = { "2", "4", "5", "6", "7", "8", "9", "11", "12", "13", "14", "15", "16" };
        private readonly string[] _columnsTable7 = { "2", "4", "5", "6", "7", "8", "9", "11", "12", "13", "14", "15", "16" };
        private readonly string[] _columnsTable8 = { "2", "5" };
        private readonly string[] _columnsTable9 = { "2", "7", "9" };
        private readonly string[] _columnsTableLetal1 = { "2", "3","4","5","6","7" };

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
                "Таблица 6" => FillTable6(form),
                "Таблица 7" => FillTable6(form),
                "Таблица 8" => FillTable1(form),
                "Таблица 9" => FillTable4(form),
                "Таблица 1Л" => FillTableLetal(form),
                "Таблица 2Л" => FillTableLetal(form),
                _ => FillTable5()
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
               
                //startRow = 16;
                rowDataIndexFirst = "7";
                rowDataIndexSecond = "9";
                dictionary = FindColumnIndexies(_columnsTable9, startRow - 1);
               
                for (int i = startRow; i <= lastRow; i++)
                {
                    var data = new ReportZpz2025DataDto
                    {
                        Code = ObjWorkSheet.Cells[i, dictionary["2"]].Value,
                        CountSmo = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary[rowDataIndexFirst]].Value),
                        CountSmoAnother = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary[rowDataIndexSecond]].Value) 
                    };
                    list.Add(data); 
                }
            }

            return list.ToArray();
        }

        private ReportZpz2025DataDto[] FillTable5()
        {
            var list = new List<ReportZpz2025DataDto>();
            int countWorkSheet = ObjWorkBook.Worksheets.Count;

            for (int currentList = 1; currentList <= countWorkSheet; currentList++)
            {
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[currentList];
                int lastRow = GetLastRow();
                int startRow = GetStartRow(); //currentList == 1 ? 16 : 5;
                Dictionary<string, int> dictionary = FindColumnIndexies(_columnsTable5A, startRow - 1);

                for (int i = startRow; i <= lastRow; i++)
                {
                    var data = new ReportZpz2025DataDto
                    {
                        Code = ObjWorkSheet.Cells[i, dictionary["2"]].Value,
                        CountOutOfSmo = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["4"]].Value),
                        CountAmbulatory = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["5"]].Value),
                        CountDs = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["6"]].Value),
                        CountDsVmp = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["7"]].Value),
                        CountStac = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["8"]].Value),
                        CountStacVmp = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["9"]].Value)
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
                int startRow = GetStartRow();

                //startRow = currentList == 1 ? 15 : 4;
                dictionary = FindColumnIndexies(_columnsTable8, startRow - 1);


                for (int i = startRow; i <= lastRow; i++)
                {
                    var data = new ReportZpz2025DataDto
                    {
                        Code = ObjWorkSheet.Cells[i, dictionary["2"]].Value,
                        CountSmo = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["5"]].Value)
                    };
                    list.Add(data);
                }
            }

            return list.ToArray();
        }

        private ReportZpz2025DataDto[] FillTable6(string form)
        {
            var list = new List<ReportZpz2025DataDto>();
            int countWorkSheet = ObjWorkBook.Worksheets.Count;

            for (int currentList = 1; currentList <= countWorkSheet; currentList++)
            {
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[currentList];
                int lastRow = GetLastRow();
                int startRow = GetStartRow(); //currentList == 1 ? 16 : 5;
                Dictionary<string, int> dictionary = form == "Таблица 2" ?
                   FindColumnIndexies(_columnsTable6, startRow - 1) :
                   FindColumnIndexies(_columnsTable7, startRow - 1);

                for (int i = startRow; i <= lastRow; i++)
                {
                    var data = new ReportZpz2025DataDto
                    {
                        Code = ObjWorkSheet.Cells[i, dictionary["2"]].Value,
                        CountOutOfSmo = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["4"]].Value),
                        CountAmbulatory = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["5"]].Value),
                        CountDs = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["6"]].Value),
                        CountDsVmp = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["7"]].Value),
                        CountStac = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["8"]].Value),
                        CountStacVmp = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["9"]].Value),
                        CountOutOfSmoAnother = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["11"]].Value),
                        CountAmbulatoryAnother = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["12"]].Value),
                        CountDsAnother = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["13"]].Value),
                        CountDsVmpAnother = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["14"]].Value),
                        CountStacAnother = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["15"]].Value),
                        CountStacVmpAnother = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["16"]].Value),
                    };
                    list.Add(data);
                }
            }

            return list.ToArray();
        }

        private ReportZpz2025DataDto[] FillTableLetal(string form)
        {
            var list = new List<ReportZpz2025DataDto>();
            int countWorkSheet = ObjWorkBook.Worksheets.Count;

            for (int currentList = 1; currentList <= countWorkSheet; currentList++)
            {
                ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[currentList];
                int lastRow = GetLastRow();

                Dictionary<string, int> dictionary = new Dictionary<string, int>();
                int startRow = 0;

                switch (form)
                {
                    case "Таблица 1Л":
                        startRow = GetStartRow();
                        lastRow = 35;
                        dictionary = FindColumnIndexies(_columnsTableLetal1, startRow - 1);
                        break;
                    case "Таблица 2Л":
                        startRow = GetStartRow();
                        lastRow = 45;
                        dictionary = FindColumnIndexies(_columnsTableLetal1, startRow - 1);
                        break;

                }

                for (int i = startRow; i <= lastRow; i++)
                {
                    var data = new ReportZpz2025DataDto
                    {
                        Code = ObjWorkSheet.Cells[i, dictionary["2"]].Value,                    
                        CountAmbulatory = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["3"]].Value),
                        CountStac = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["4"]].Value),
                        CountDs = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["5"]].Value),
                        CountOutOfSmoAnother = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["6"]].Value),
                        CountSmo = GlobalUtils.TryParseDecimal(ObjWorkSheet.Cells[i, dictionary["7"]].Value),
                      
                    };
                    list.Add(data);
                }
            }

            return list.ToArray();
        }

    }
}