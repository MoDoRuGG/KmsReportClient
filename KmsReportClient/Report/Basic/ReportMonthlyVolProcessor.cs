using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using KmsReportClient.Excel.Creator.Base;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.XML;
using KmsReportClient.Support;
using NLog;

namespace KmsReportClient.Report.Basic
{
    class ReportMonthlyVolProcessor : AbstractReportProcessor<ReportMonthlyVol>
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        private readonly string[] _forms = { "Стационарная помощь", "Дневной стационар", "АПП", "Скорая медицинская помощь" };



        private readonly string[][] _headers = {
            new[]
            { "Всего случаев в реестре",
              "Принято к оплате случаев",
              "План для СМО по МЭЭ, количество случаев по плану",
              "План для СМО по МЭЭ, % по плану",
              "Выполнено СМО по МЭЭ, количество случаев экспертиз",
              "Выполнено СМО по МЭЭ, % выполнения",
              "План для СМО по ЭКМП, количество случаев по плану",
              "План для СМО по ЭКМП, % по плану",
              "Выполнено СМО по ЭКМП, количество случаев экспертиз",
              "Выполнено СМО по ЭКМП, % выполнения",
              },
            new[]
            { "Всего случаев в реестре",
              "Принято к оплате случаев",
              "План для СМО по МЭЭ, количество случаев по плану",
              "План для СМО по МЭЭ, % по плану",
              "Выполнено СМО по МЭЭ, количество случаев экспертиз",
              "Выполнено СМО по МЭЭ, % выполнения",
              "План для СМО по ЭКМП, количество случаев по плану",
              "План для СМО по ЭКМП, % по плану",
              "Выполнено СМО по ЭКМП, количество случаев экспертиз",
              "Выполнено СМО по ЭКМП, % выполнения",
              },
            new[]
            { "Всего случаев в реестре",
              "Принято к оплате случаев",
              "План для СМО по МЭЭ, количество случаев по плану",
              "План для СМО по МЭЭ, % по плану",
              "Выполнено СМО по МЭЭ, количество случаев экспертиз",
              "Выполнено СМО по МЭЭ, % выполнения",
              "План для СМО по ЭКМП, количество случаев по плану",
              "План для СМО по ЭКМП, % по плану",
              "Выполнено СМО по ЭКМП, количество случаев экспертиз",
              "Выполнено СМО по ЭКМП, % выполнения",
              },
            new[]
            { "Всего случаев в реестре",
              "Принято к оплате случаев",
              "План для СМО по МЭЭ, количество случаев по плану",
              "План для СМО по МЭЭ, % по плану",
              "Выполнено СМО по МЭЭ, количество случаев экспертиз",
              "Выполнено СМО по МЭЭ, % выполнения",
              "План для СМО по ЭКМП, количество случаев по плану",
              "План для СМО по ЭКМП, % по плану",
              "Выполнено СМО по ЭКМП, количество случаев экспертиз",
              "Выполнено СМО по ЭКМП, % выполнения",
              },
        };

        private readonly Dictionary<string, string> _headersMap = new Dictionary<string, string>
        {
            { "Стационарная помощь", "Период экспертизы" },
            { "Дневной стационар", "Период экспертизы" },
            { "АПП", "Период экспертизы" },
            { "Скорая медицинская помощь", "Период экспертизы" },    
        };

        public ReportMonthlyVolProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
            base(inClient, dgv, cmb, txtb, page,
                XmlFormTemplate.MonthlyVol.GetDescription(),
                Log,
                ReportGlobalConst.ReportMonthlyVol,
                reportsDictionary)
        {
            InitReport();
        }

        public override void InitReport()
        {
            Report = new ReportMonthlyVol { ReportDataList = new ReportMonthlyVolDto[ThemesList.Count], IdType = IdReportType };

            int i = 0;
            foreach (var theme in ThemesList.Select(x => x.Key))
            {
                Report.ReportDataList[i++] = new ReportMonthlyVolDto { Theme = theme };
            }
        }

        public override AbstractReport CollectReportFromWs(string yymm)
        {
            var request = new GetReportRequest
            {
                Body = new GetReportRequestBody
                {
                    filialCode = FilialCode,
                    yymm = yymm,
                    reportType = ReportType.MonthlyVol
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response as ReportMonthlyVol;
        }

        public override void MapForAutoFill(AbstractReport report)
        {
            if (report == null)
            {
                return;
            }
            var inReport = report as ReportMonthlyVol;

            var index = Report.ReportDataList.ToList().FindIndex(x => x.Theme == Cmb.Text);
            var inTheme = inReport.ReportDataList.Single(x => x.Theme == Cmb.Text);
            Report.ReportDataList[index] = inTheme;
        }

        public override void FillDataGridView(string form)
        {
            if (form == null)
            {
                return;
            }
            if (_forms.Contains(form))
            {
                FillDgvForms(Dgv, form);
            }
        }

        protected override void FillReport(string form)
        {
            if (form == null)
            {
                return;
            }
            if (_forms.Contains(form))
            {
                FillThemesForms(Dgv, form);
            }
        }

        
        public override bool IsVisibleBtnDownloadExcel() => true;

        public override bool IsVisibleBtnHandle() => true;
        public override bool IsVisibleBtnSummary() => false;

        public override string ValidReport()
        {
            string message = "";
            //string[] validForms = { "Таблица 6", "Таблица 7" };
            //foreach (var data in Report.ReportDataList.Where(x => validForms.Contains(x.Theme)))
            //{
            //    if (data.Data == null)
            //    {
            //        continue;
            //    }

            //    string localMessage = "";
            //    string lastSumRow = data.Theme == "Таблица 6" ? "2.5, 2.6" : "2.5";

            //    decimal gr4 = data.Data.Where(x => x.Code == "2").Sum(x => x.CountOutOfSmo);
            //    decimal gr4Another = data.Data.Where(x => x.Code.StartsWith("2") && x.Code.Length == 3).Sum(x => x.CountOutOfSmo);
            //    decimal gr5 = data.Data.Where(x => x.Code == "2").Sum(x => x.CountAmbulatory);
            //    decimal gr5Another = data.Data.Where(x => x.Code.StartsWith("2") && x.Code.Length == 3).Sum(x => x.CountAmbulatory);
            //    decimal gr6 = data.Data.Where(x => x.Code == "2").Sum(x => x.CountDs);
            //    decimal gr6Another = data.Data.Where(x => x.Code.StartsWith("2") && x.Code.Length == 3).Sum(x => x.CountDs);
            //    decimal gr7 = data.Data.Where(x => x.Code == "2").Sum(x => x.CountDsVmp);
            //    decimal gr7Another = data.Data.Where(x => x.Code.StartsWith("2") && x.Code.Length == 3).Sum(x => x.CountDsVmp);
            //    decimal gr8 = data.Data.Where(x => x.Code == "2").Sum(x => x.CountStac);
            //    decimal gr8Another = data.Data.Where(x => x.Code.StartsWith("2") && x.Code.Length == 3).Sum(x => x.CountStac);
            //    decimal gr9 = data.Data.Where(x => x.Code == "2").Sum(x => x.CountStacVmp);
            //    decimal gr9Another = data.Data.Where(x => x.Code.StartsWith("2") && x.Code.Length == 3).Sum(x => x.CountStacVmp);
            //    decimal gr11 = data.Data.Where(x => x.Code == "2").Sum(x => x.CountOutOfSmoAnother);
            //    decimal gr11Another = data.Data.Where(x => x.Code.StartsWith("2") && x.Code.Length == 3).Sum(x => x.CountOutOfSmoAnother);
            //    decimal gr12 = data.Data.Where(x => x.Code == "2").Sum(x => x.CountAmbulatoryAnother);
            //    decimal gr12Another = data.Data.Where(x => x.Code.StartsWith("2") && x.Code.Length == 3).Sum(x => x.CountAmbulatoryAnother);
            //    decimal gr13 = data.Data.Where(x => x.Code == "2").Sum(x => x.CountDsAnother);
            //    decimal gr13Another = data.Data.Where(x => x.Code.StartsWith("2") && x.Code.Length == 3).Sum(x => x.CountDsAnother);
            //    decimal gr14 = data.Data.Where(x => x.Code == "2").Sum(x => x.CountDsVmpAnother);
            //    decimal gr14Another = data.Data.Where(x => x.Code.StartsWith("2") && x.Code.Length == 3).Sum(x => x.CountDsVmpAnother);
            //    decimal gr15 = data.Data.Where(x => x.Code == "2").Sum(x => x.CountStacAnother);
            //    decimal gr15Another = data.Data.Where(x => x.Code.StartsWith("2") && x.Code.Length == 3).Sum(x => x.CountStacAnother);
            //    decimal gr16 = data.Data.Where(x => x.Code == "2").Sum(x => x.CountStacVmpAnother);
            //    decimal gr16Another = data.Data.Where(x => x.Code.StartsWith("2") && x.Code.Length == 3).Sum(x => x.CountStacVmpAnother);
            //    if (gr4 < gr4Another)
            //    {
            //        localMessage += $"гр.4 - значение стр.2 должно быть больше или равна сумме строк 2.1, 2.2, 2.3, 2.4, {lastSumRow} \r\n";
            //    }

            //    if (gr5 < gr5Another)
            //    {
            //        localMessage += $"гр.5 - значение стр.2 должно быть больше или равна сумме строк 2.1, 2.2, 2.3, 2.4, {lastSumRow} \r\n";
            //    }

            //    if (gr6 < gr6Another)
            //    {
            //        localMessage += $"гр.6 - значение стр.2 должно быть больше или равна сумме строк 2.1, 2.2, 2.3, 2.4, {lastSumRow} \r\n";
            //    }

            //    if (gr7 < gr7Another)
            //    {
            //        localMessage += $"гр.7 - значение стр.2 должно быть больше или равна сумме строк 2.1, 2.2, 2.3, 2.4, {lastSumRow} \r\n";
            //    }

            //    if (gr8 < gr8Another)
            //    {
            //        localMessage += $"гр.8 - значение стр.2 должно быть больше или равна сумме строк 2.1, 2.2, 2.3, 2.4, {lastSumRow} \r\n";
            //    }

            //    if (gr9 < gr9Another)
            //    {
            //        localMessage += $"гр.9 - значение стр.2 должно быть больше или равна сумме строк 2.1, 2.2, 2.3, 2.4, {lastSumRow} \r\n";
            //    }

            //    if (gr11 < gr11Another)
            //    {
            //        localMessage += $"гр.11 - значение стр.2 должно быть больше или равна сумме строк 2.1, 2.2, 2.3, 2.4, {lastSumRow} \r\n";
            //    }

            //    if (gr12 < gr12Another)
            //    {
            //        localMessage += $"гр.12 - значение стр.2 должно быть больше или равна сумме строк 2.1, 2.2, 2.3, 2.4, {lastSumRow} \r\n";
            //    }

            //    if (gr13 < gr13Another)
            //    {
            //        localMessage += $"гр.13 - значение стр.2 должно быть больше или равна сумме строк 2.1, 2.2, 2.3, 2.4, {lastSumRow} \r\n";
            //    }

            //    if (gr14 < gr14Another)
            //    {
            //        localMessage += $"гр.14 - значение стр.2 должно быть больше или равна сумме строк 2.1, 2.2, 2.3, 2.4, {lastSumRow} \r\n";
            //    }

            //    if (gr15 < gr15Another)
            //    {
            //        localMessage += $"гр.15 - значение стр.2 должно быть больше или равна сумме строк 2.1, 2.2, 2.3, 2.4, {lastSumRow} \r\n";
            //    }

            //    if (gr16 < gr16Another)
            //    {
            //        localMessage += $"гр.16 - значение стр.2 должно быть больше или равна сумме строк 2.1, 2.2, 2.3, 2.4, {lastSumRow} \r\n";
            //    }

            //    if (localMessage.Length > 0)
            //    {
            //        message += $"{data.Theme}. \r\n {localMessage}";
            //    }
            //}
            //if (message.Length > 0)
            //{
            //    message = "Форма ЗПЗ. " + Environment.NewLine + message;
            //}
            return message;
        }

        public override void ToExcel(string filename, string filialName)
        {
            //var mm = YymmUtils.GetMonth(Report.Yymm.Substring(2, 2)) + " 20" + Report.Yymm.Substring(0, 2);
            //var excel = new ExcelMonthlyVolCreator(filename, ExcelForm.MonthlyVol, mm, filialName);
            //excel.CreateReport(Report, null);
        }

        public override void SaveToDb()
        {
            var request = new SaveReportRequest
            {
                Body = new SaveReportRequestBody
                
                {
                    filialCode = CurrentUser.FilialCode,
                    idUser = CurrentUser.IdUser,
                    report = Report,
                    yymm = Report.Yymm,
                    reportType = ReportType.MonthlyVol
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportMonthlyVol;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
            Report.DataSource = response.DataSource;
            
        }

        public void SetFormula() { }

        public override void SaveReportDataSourceExcel()
        {
            var request = new SaveReportDataSourceExcelRequest
            {
                Body = new SaveReportDataSourceExcelRequestBody

                {
                    report = Report,
                    filialCode = CurrentUser.FilialCode,
                    idUser = CurrentUser.IdUser,
                    yymm = Report.Yymm,
                    reportType = ReportType.MonthlyVol
                }
            };
            var response = Client.SaveReportDataSourceExcel(request).Body.SaveReportDataSourceExcelResult as ReportMonthlyVol;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
            Report.DataSource = response.DataSource;

        }

        public override void SaveReportDataSourceHandle()
        {
            var request = new SaveReportDataSourceHandleRequest
            {
                Body = new SaveReportDataSourceHandleRequestBody

                {
                    report = Report,
                    filialCode = CurrentUser.FilialCode,
                    idUser = CurrentUser.IdUser,
                    yymm = Report.Yymm,
                    reportType = ReportType.MonthlyVol
                }
            };
            var response = Client.SaveReportDataSourceHandle(request).Body.SaveReportDataSourceHandleResult as ReportMonthlyVol;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
            Report.DataSource = response.DataSource;

        }

        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource)
        {
            var array = new ArrayOfString();
            array.AddRange(filialList);
            var request = new CollectSummaryReportRequest
            {
                Body = new CollectSummaryReportRequestBody
                {
                    filials = array,
                    status = status,
                    yymmStart = yymmStart,
                    yymmEnd = yymmEnd,
                    reportType = ReportType.MonthlyVol
                }
            };
            var response = Client.CollectSummaryReport(request);
            Report = response.Body.CollectSummaryReportResult as ReportMonthlyVol;
            Report.IdType = IdReportType;
            Report.Yymm = yymmEnd;
        }

        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            var formsList = ThemesList.Select(x => x.Key).OrderBy(x => x).ToList();
            var index = formsList.IndexOf(form);
            var currentHeaders = _headers[index];
            CreateDgvColumnsForTheme(Dgv, 400, _headersMap[form], currentHeaders);

            int countRows = ThemeTextData.Tables_fromxml.Single(x => x.TableName_fromxml == form).RowsCount_fromxml;
            foreach (var row in table)
            {
                var dgvRow = new DataGridViewRow();
                var cellName = new DataGridViewTextBoxCell
                {
                    Value = row.RowText_fromxml
                };
                var cellNum = new DataGridViewTextBoxCell
                {
                    Value = row.RowNum_fromxml
                };
                dgvRow.Cells.Add(cellName);
                dgvRow.Cells.Add(cellNum);

                var exclusionCells = row.ExclusionCells_fromxml?.Split(',');
                for (int i = 2; i < countRows; i++)
                {
                    bool isNeedExcludeSum = exclusionCells?.Contains(i.ToString()) ?? false;
                    var cell = new DataGridViewTextBoxCell
                    {
                        Value = row.Exclusion_fromxml || isNeedExcludeSum ? "x" : "0"
                    };
                    dgvRow.Cells.Add(cell);

                    if (isNeedExcludeSum)
                    {
                        cell.ReadOnly = true;
                        cell.Style.BackColor = Color.DarkGray;
                    }
                }
                int rowIndex = Dgv.Rows.Add(dgvRow);
                if (row.Exclusion_fromxml)
                {
                    Dgv.Rows[rowIndex].ReadOnly = true;
                    Dgv.Rows[rowIndex].DefaultCellStyle.BackColor = Color.LightCyan;
                }

            }
        }

        private void CreateDgvColumnsForTheme(DataGridView dgvReport, int widthFirstColumn, string mainHeader,
            string[] columns)
        {
            CreateDgvCommonColumns(dgvReport, widthFirstColumn, mainHeader);
            foreach (var column in columns)
            {
                var dgvColumn = new DataGridViewTextBoxColumn
                {
                    HeaderText = column,
                    Width = 100,
                    ReadOnly = false,
                    SortMode = DataGridViewColumnSortMode.NotSortable
                };
                dgvReport.Columns.Add(dgvColumn);
            }
        }

        private void CreateDgvCommonColumns(DataGridView dgvReport, int widthFirstColumn, string mainHeader)
        {
            dgvReport.AllowUserToAddRows = false;
            dgvReport.ColumnHeadersVisible = true;
            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = mainHeader,
                Width = widthFirstColumn,
                DataPropertyName = "Indicator",
                Name = "Indicator",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.Azure
                }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "№ строки",
                Width = 50,
                DataPropertyName = "NumRow",
                Name = "NumRow",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.Azure
                }
            };
            dgvReport.Columns.Add(column);

        }


        private void FillThemesForms(DataGridView dgvReport, string form)
        {
            var reportDto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportDto == null)
            {
                return;
            }

            reportDto.Data = (from DataGridViewRow row in dgvReport.Rows
                              let rowNum = row.Cells[1].Value.ToString().Trim()
                              where !IsNotNeedFillRow(form, rowNum)
                              select new ReportMonthlyVolDataDto
                              {
                                  Code = rowNum,
                                  CountSluch = GlobalUtils.TryParseInt(row.Cells[2].Value),
                                  CountAppliedSluch = GlobalUtils.TryParseInt(row.Cells[3].Value),
                                  CountSluchMEE = GlobalUtils.TryParseInt(row.Cells[4].Value),
                                  CountSluchEKMP = GlobalUtils.TryParseInt(row.Cells[5].Value)
                              }).ToArray();
        }

        
        private void FillDgvForms(DataGridView dgvReport, string form)
        {
            var reportDto = Report.ReportDataList?.Single(x => x.Theme == form);
            if (reportDto?.Data == null || reportDto.Data.Length == 0)
            {
                return;
            }

            var rows = ThemeTextData.Tables_fromxml.Where(x => x.TableName_fromxml == form).SelectMany(x => x.Rows_fromxml).ToList();
            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                var rowNum = row.Cells[1].Value.ToString().Trim();
                bool exclusionsRow = rows.Single(x => x.RowNum_fromxml == rowNum).Exclusion_fromxml;

                var data = reportDto.Data.SingleOrDefault(x => x.Code == rowNum);
                if (data != null)
                {
                    row.Cells[2].Value = data.CountSluch.ToString().Replace(",00", "");
                    row.Cells[3].Value = data.CountAppliedSluch.ToString().Replace(",00", "");
                    row.Cells[4].Value = data.CountSluchMEE.ToString().Replace(",00", "");
                    row.Cells[5].Value = data.CountSluchEKMP.ToString().Replace(",00", "");
                }
            }
        }
    }
}
