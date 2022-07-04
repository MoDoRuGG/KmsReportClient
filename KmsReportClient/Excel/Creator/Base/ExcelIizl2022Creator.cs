using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.Excel;
using KmsReportClient.Support;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Base
{
    public class ExcelIizl2022Creator : ExcelBaseCreator<ReportIizl>
    {

        public ExcelIizl2022Creator(
          string filename,
          ExcelForm reportName,
          string header,
          string filialName) : base(filename, reportName, header, filialName, false) { }

        private readonly List<ReportDictionary> _electronicMeansThemesDic = new List<ReportDictionary> {
            new ReportDictionary {TableName = "Тема Д1-Э", StartRow = 18},
            new ReportDictionary {TableName = "Тема Д2-Э", StartRow = 23},
            new ReportDictionary {TableName = "Тема Д3-Э", StartRow = 28},
            new ReportDictionary {TableName = "Тема Д4-Э", StartRow = 33},
            new ReportDictionary {TableName = "Тема П-Э", StartRow = 38},
            new ReportDictionary {TableName = "Тема С-Э", StartRow = 43},
            new ReportDictionary {TableName = "Тема К-Э", StartRow = 46},
            new ReportDictionary {TableName = "Тема О-Э", StartRow = 49},
            new ReportDictionary {TableName = "Тема УД-Э", StartRow = 54},
            new ReportDictionary {TableName = "Тема И-Э", StartRow = 59}
        };

        private readonly List<ReportDictionary> _writtenInformationThemesDic = new List<ReportDictionary> {
            new ReportDictionary {TableName = "Тема Д1-П", StartRow = 23},
            new ReportDictionary {TableName = "Тема Д2-П", StartRow = 33},
            new ReportDictionary {TableName = "Тема Д3-П", StartRow = 43},
            new ReportDictionary {TableName = "Тема Д4-П", StartRow = 53},
            new ReportDictionary {TableName = "Тема П-П", StartRow = 63},
            new ReportDictionary {TableName = "Тема С-П", StartRow = 73},
            new ReportDictionary {TableName = "Тема К-П", StartRow = 78},
            new ReportDictionary {TableName = "Тема О-П", StartRow = 83},
            new ReportDictionary {TableName = "Тема УД-П", StartRow = 91},
            new ReportDictionary {TableName = "Тема И-П", StartRow = 101}
        };



        protected override void FillReport(ReportIizl report, ReportIizl yearReport)
        {
            FillElectronicMeansThemes(report);
            FillWrittenInformationThemes(report);
            FillAgreement(report);

            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[4];

            ObjWorkSheet.Cells[5, 4] = FilialName;
            ObjWorkSheet.Cells[6, 4] = Header;
        }

        private void FillAgreement(ReportIizl report)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[3];
            var themeData = report.ReportDataList.FirstOrDefault(x => x.Theme.Contains("Соглас"));
            int rowIndex = 4;
            if (themeData != null)
            {
                foreach (var row in themeData.Data.OrderBy(x => x.Code))
                {
                    ObjWorkSheet.Cells[rowIndex++, 7] = row.CountPersFirst;
                }
            }

        }

        private void FillWrittenInformationThemes(ReportIizl report)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[2];

            ObjWorkSheet.Cells[5, 4] = FilialName;
            ObjWorkSheet.Cells[6, 4] = Header;
            ObjWorkSheet.Cells[111, 2] = CurrentUser.Director;
            ObjWorkSheet.Cells[111, 6] = DateTime.Today;

            var writtenInformationThemes = report.ReportDataList.Where(x => x.Theme.EndsWith("-П"));
            int rowIndex;
            foreach (var theme in _writtenInformationThemesDic)
            {
                rowIndex = theme.StartRow;
                var currentTheme = writtenInformationThemes.FirstOrDefault(x => x.Theme == theme.TableName);
                if (currentTheme != null)
                {
                    foreach (var row in currentTheme.Data)
                    {
                        if (ObjWorkSheet.Cells[rowIndex, 1].Text == "Сумма")
                        {
                            continue;
                        }

                        if (ObjWorkSheet.Cells[rowIndex, 3].Text != "х")
                        {
                            ObjWorkSheet.Cells[rowIndex, 3] = row.CountPersFirst;
                        }

                        if (ObjWorkSheet.Cells[rowIndex, 4].Text != "х")
                        {
                            ObjWorkSheet.Cells[rowIndex, 4] = row.CountPersRepeat;
                        }

                        if (ObjWorkSheet.Cells[rowIndex, 6].Text != "х")
                        {
                            ObjWorkSheet.Cells[rowIndex, 6] = row.TotalCost;
                        }

                        if (ObjWorkSheet.Cells[rowIndex, 8].Text != "х")
                        {
                            ObjWorkSheet.Cells[rowIndex, 8] = row.AccountingDocument;
                        }

                        rowIndex++;

                    }
                }
            }
        }


        private void FillElectronicMeansThemes(ReportIizl report)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];

            ObjWorkSheet.Cells[5, 4] = FilialName;
            ObjWorkSheet.Cells[6, 4] = Header;
            ObjWorkSheet.Cells[64, 2] = CurrentUser.Director;
            ObjWorkSheet.Cells[64, 8] = DateTime.Today;

            var electronicMeansThemes = report.ReportDataList.Where(x => x.Theme.EndsWith("-Э"));
            int rowIndex;
            foreach (var theme in _electronicMeansThemesDic)
            {
                rowIndex = theme.StartRow;
                var currentTheme = electronicMeansThemes.FirstOrDefault(x => x.Theme == theme.TableName);
                if (currentTheme != null)
                {
                    foreach (var row in currentTheme.Data)
                    {
                        if (ObjWorkSheet.Cells[rowIndex, 1].Text == "Сумма")
                        {
                            continue;
                        }

                        if (ObjWorkSheet.Cells[rowIndex, 3].Text != "х")
                        {
                            ObjWorkSheet.Cells[rowIndex, 3] = row.CountPersFirst;
                        }

                        if (ObjWorkSheet.Cells[rowIndex, 4].Text != "х")
                        {
                            ObjWorkSheet.Cells[rowIndex, 4] = row.CountPersRepeat;
                        }

                        if (ObjWorkSheet.Cells[rowIndex, 7].Text != "х")
                        {
                            ObjWorkSheet.Cells[rowIndex, 7] = row.CountMessages;
                        }

                        if (ObjWorkSheet.Cells[rowIndex, 8].Text != "х")
                        {
                            ObjWorkSheet.Cells[rowIndex, 8] = row.TotalCost;
                        }

                        if (theme.TableName.Contains("Тема С-Э") || theme.TableName.Contains("Тема К-Э"))
                        {
                            if (ObjWorkSheet.Cells[rowIndex, 9].Text != "х")
                            {
                                ObjWorkSheet.Cells[rowIndex, 9] = row.AverageCostPerMessage;
                            }

                            if (ObjWorkSheet.Cells[rowIndex, 10].Text != "х")
                            {
                                ObjWorkSheet.Cells[rowIndex, 10] = row.AverageCostOfInforming1PL;
                            }
                        }

                        if (ObjWorkSheet.Cells[rowIndex, 11].Text != "х")
                        {
                            ObjWorkSheet.Cells[rowIndex, 11] = row.AccountingDocument;
                        }

                        rowIndex++;

                    }
                }
            }
        }
    }
}
