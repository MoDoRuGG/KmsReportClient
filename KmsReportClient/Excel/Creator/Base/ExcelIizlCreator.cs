using System;
using System.Collections.Generic;
using System.Linq;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.Excel;

namespace KmsReportClient.Excel.Creator.Base
{
    class ExcelIizlCreator : ExcelBaseCreator<ReportIizl>
    {
        private readonly List<ReportDictionary> _iizlDictionaries = new List<ReportDictionary> {
            new ReportDictionary {TableName = "Согласие", StartRow = 29},
            new ReportDictionary {TableName = "Тема Д1", StartRow = 43},
            new ReportDictionary {TableName = "Тема Д2", StartRow = 62},
            new ReportDictionary {TableName = "Тема Д3", StartRow = 81},
            new ReportDictionary {TableName = "Тема Д4", StartRow = 100},
            new ReportDictionary {TableName = "Тема П", StartRow = 119},
            new ReportDictionary {TableName = "Тема С", StartRow = 138},
            new ReportDictionary {TableName = "Тема К", StartRow = 150},
            new ReportDictionary {TableName = "Тема О", StartRow = 162},
            new ReportDictionary {TableName = "Тема Кор", StartRow = 180}
        };

        public ExcelIizlCreator(
            string filename,
            ExcelForm reportName,
            string header,
            string filialName) : base(filename, reportName, header, filialName, false) { }

        protected override void FillReport(ReportIizl report, ReportIizl yearReport)
        {
            foreach (var table in _iizlDictionaries)
            {
                var themeData = report.ReportDataList.SingleOrDefault(x => x.Theme.ToLower() == table.TableName.ToLower());
                if (themeData == null)
                {
                    return;
                }
                int currentIndex = table.StartRow;

                if (themeData.Theme.StartsWith("Тема"))
                {
                    string prefix = themeData.Theme.Split(' ')[1];
                    string[] suffixes = { "У", "П" };
                 
                    foreach (var suffix in suffixes)
                    {
                        var dataList = themeData.Data?.Where(x => x.Code.ToLower().StartsWith($"{prefix.ToLower()}-{suffix.ToLower()}")).OrderBy(x => x.Code);
                       
                         //Console.WriteLine($"{dataList.Count()} количество");
                         //Console.WriteLine($"{prefix.ToLower()}-{suffix.ToLower()}");
                    



                        foreach (var data in dataList.ToList())
                        {
                            
                            ObjWorkSheet.Cells[currentIndex, 3] = data.CountPersFirst;
                            ObjWorkSheet.Cells[currentIndex, 4] = data.CountPersRepeat;
                            ObjWorkSheet.Cells[currentIndex, 5] = data.CountMessages;
                            ObjWorkSheet.Cells[currentIndex, 6] = data.TotalCost;
                            ObjWorkSheet.Cells[currentIndex++, 7] = data.AccountingDocument;
                        }

                        currentIndex++;
                    }

                    ObjWorkSheet.Cells[currentIndex, 3] = themeData.TotalPersFirst;
                    ObjWorkSheet.Cells[currentIndex, 4] = themeData.TotalPersRepeat;
                }
                else
                {
                    foreach (var data in themeData.Data.OrderBy(x => x.Code))
                    {
                        ObjWorkSheet.Cells[currentIndex++, 7] = data.CountPersFirst;
                    }
                }
            }

            ObjWorkSheet.Cells[5, 4] = FilialName;
            ObjWorkSheet.Cells[6, 4] = Header;
            ObjWorkSheet.Cells[199, 2] = CurrentUser.Director;
            ObjWorkSheet.Cells[199, 6] = DateTime.Today;
        }
    }
}
