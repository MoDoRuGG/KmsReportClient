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
    public class ReportIizlProcessor : AbstractReportProcessor<ReportIizl>
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        public ReportIizlProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv,
            ComboBox cmb, TextBox txtb, TabPage page) :
            base(inClient, dgv, cmb, txtb, page,
                XmlFormTemplate.Iizl.GetDescription(),
                Log,
                ReportGlobalConst.ReportIizl,
                reportsDictionary)
        {
            InitReport();
        }

        public override void InitReport()
        {
            Report = new ReportIizl {ReportDataList = new ReportIizlDto[ThemesList.Count], IdType = IdReportType};

            int i = 0;
            foreach (var theme in ThemesList.Select(x => x.Key))
            {
                var themeData = ThemeTextData.Tables_fromxml.Single(x => x.TableName_fromxml == theme);
                var rows = themeData.Rows_fromxml.Select(x => new ReportIizlDataDto {Code = x.RowNum_fromxml}).ToArray();

                Report.ReportDataList[i++] = new ReportIizlDto {Theme = theme, Data = rows};
            }
        }

        public override AbstractReport CollectReportFromWs(string yymm)
        {
            var request = new GetReportRequest {
                Body = new GetReportRequestBody {filialCode = FilialCode, yymm = yymm, reportType = ReportType.Iizl}
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as ReportIizl;
        }

        public override void MapForAutoFill(AbstractReport report)
        {
            if (report == null)
            {
                return;
            }

            var inReport = report as ReportIizl;

            var index = Report.ReportDataList.ToList().FindIndex(x => x.Theme == Cmb.Text);
            var inTheme = inReport.ReportDataList.Single(x => x.Theme == Cmb.Text);
            Report.ReportDataList[index] = inTheme;
        }

        protected override void FillReport(string form)
        {
            if (form == null)
            {
                return;
            }

            if (form.StartsWith("Тема"))
            {
                FillThemesReport(Dgv, form);
            }
            else
            {
                FillInfAgreeReport(Dgv, form);
            }
        }

        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            if (form.StartsWith("Тема"))
            {
                CreateDgvThemeColumns(Dgv);
            }
            else
            {
                CreateDgvInfAgreeColumns(Dgv);
            }

            int countRows = ThemeTextData.Tables_fromxml.Single(x => x.TableName_fromxml == form).RowsCount_fromxml;
            foreach (var row in table)
            {
                var dgvRow = new DataGridViewRow();
                var cellName = new DataGridViewTextBoxCell {Value = row.RowText_fromxml};
                var cellNum = new DataGridViewTextBoxCell {Value = row.RowNum_fromxml};
                dgvRow.Cells.Add(cellName);
                dgvRow.Cells.Add(cellNum);

                var exclusionCells = row.ExclusionCells_fromxml?.Split(',');
                for (int i = 2; i < countRows; i++)
                {
                    bool isNeedExcludeSum = exclusionCells?.Contains(i.ToString()) ?? false;
                    var cell = new DataGridViewTextBoxCell {Value = isNeedExcludeSum ? "x" : "0"};
                    dgvRow.Cells.Add(cell);

                    if (isNeedExcludeSum)
                    {
                        cell.ReadOnly = true;
                        cell.Style.BackColor = Color.DarkGray;
                    }
                }

                Dgv.Rows.Add(dgvRow);
            }
        }
        public override void SaveReportDataSourceExcel()
        { }
        public override void SaveReportDataSourceHandle()
        { }

        private void FillInfAgreeReport(DataGridView dgvReport, string form)
        {
            var reportIizlDto = Report.ReportDataList.Single(x => x.Theme == form);


            if(reportIizlDto==null)
            {
                return;
            }
            reportIizlDto.Data = (from DataGridViewRow row in dgvReport.Rows
                let code = row.Cells[1].Value.ToString()
                let sum = row.Cells[2].Value?.ToString() ?? "0"
                select new ReportIizlDataDto {Code = code, CountPersFirst = GlobalUtils.TryParseInt(sum)}).ToArray();
        }

        private void FillThemesReport(DataGridView dgvReport, string form)
        {
            var reportIizlDto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportIizlDto != null)
            {

                var reportData = new List<ReportIizlDataDto>();
                foreach (DataGridViewRow row in dgvReport.Rows)
                {
                    var code = row.Cells[1].Value.ToString();
                    var accountingDocuments = row.Cells[6].Value?.ToString() ?? "";
                    if (!code.StartsWith("И-"))
                    {
                        var data = new ReportIizlDataDto
                        {
                            Code = code,
                            CountPersFirst = GlobalUtils.TryParseInt(row.Cells[2].Value),
                            CountPersRepeat = GlobalUtils.TryParseInt(row.Cells[3].Value),
                            CountMessages = GlobalUtils.TryParseInt(row.Cells[4].Value),
                            TotalCost = GlobalUtils.TryParseDecimal(row.Cells[5].Value),
                            AccountingDocument = accountingDocuments
                        };
                        reportData.Add(data);
                    }
                    else
                    {
                        reportIizlDto.TotalPersFirst = GlobalUtils.TryParseInt(row.Cells[2].Value);
                        reportIizlDto.TotalPersRepeat = GlobalUtils.TryParseInt(row.Cells[3].Value);
                    }
                }

                reportIizlDto.Data = reportData.ToArray();
            }
        }

        public override void FillDataGridView(string form)
        {
            if (form.StartsWith("Тема"))
            {
                FillDgwThemes(Dgv, form);
            }
            else
            {
                FillDgwInfAgree(Dgv, form);
            }
        }

        public override void ToExcel(string filename, string filialName)
        {
            var mm = YymmUtils.GetMonth(Report.Yymm.Substring(2, 2)) + " 20" + Report.Yymm.Substring(0, 2);
            var excel = new ExcelIizlCreator(filename, ExcelForm.Iizl, mm, filialName);
            excel.CreateReport(Report, null);
        }

        public override void SaveToDb()
        {
            var request = new SaveReportRequest {
                Body = new SaveReportRequestBody {
                    filialCode = CurrentUser.FilialCode,
                    idUser = CurrentUser.IdUser,
                    report = Report,
                    yymm = Report.Yymm,
                    reportType = ReportType.Iizl
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportIizl;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
        }

        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource)
        {
            var array = new ArrayOfString();
            array.AddRange(filialList);
            var request = new CollectSummaryReportRequest {
                Body = new CollectSummaryReportRequestBody {
                    filials = array,
                    status = status,
                    yymmStart = yymmStart,
                    yymmEnd = yymmEnd,
                    reportType = ReportType.Iizl
                }
            };
            var response = Client.CollectSummaryReport(request);
            Report = response.Body.CollectSummaryReportResult as ReportIizl;
            Report.IdType = IdReportType;
            Report.Yymm = yymmEnd;
        }

        public override bool IsVisibleBtnDownloadExcel() => false;

        public override bool IsVisibleBtnHandle() => false;

        public override bool IsVisibleBtnSummary() => false;

        public override string ValidReport()
        {
            var message = "";
            foreach (var data in Report.ReportDataList.Where(x => !x.Theme.StartsWith("Тема")))
            {
                if (data.Data == null)
                {
                    continue;
                }

                string current = "";
                int countAll = 0;
                int countPhone = 0;
                int countSms = 0;
                int countMessengers = 0;
                int countViber = 0;
                int countPostal = 0;
                int countE = 0;
                int countEmail = 0;
                int countMobileApp = 0;
                int countDisagree = 0;
                foreach (var iizl in data.Data)
                {
                    if (iizl.Code.StartsWith("0"))
                    {
                        countAll = iizl.CountPersFirst;
                    }
                    else if (iizl.Code.StartsWith("1"))
                    {
                        countPhone = iizl.CountPersFirst;
                    }
                    else if (iizl.Code.StartsWith("2"))
                    {
                        countSms = iizl.CountPersFirst;
                    }
                    else if (iizl.Code.StartsWith("3-1"))
                    {
                        countViber = iizl.CountPersFirst;
                    }
                    else if (iizl.Code.StartsWith("3"))
                    {
                        countMessengers = iizl.CountPersFirst;
                    }
                    else if (iizl.Code.StartsWith("4"))
                    {
                        countPostal = iizl.CountPersFirst;
                    }
                    else if (iizl.Code.StartsWith("5-1"))
                    {
                        countEmail = iizl.CountPersFirst;
                    }
                    else if (iizl.Code.StartsWith("5-2"))
                    {
                        countMobileApp = iizl.CountPersFirst;
                    }
                    else if (iizl.Code.StartsWith("5"))
                    {
                        countE = iizl.CountPersFirst;
                    }
                    else if (iizl.Code.StartsWith("6"))
                    {
                        countDisagree = iizl.CountPersFirst;
                    }
                }

                var maxCountAll = countPhone + countSms + countMessengers + countPostal + countE;
                var minCountAll = Math.Max(countPhone,
                    Math.Max(countSms, Math.Max(countMessengers, Math.Max(countPostal, countE))));
                if (countAll > maxCountAll || countAll < minCountAll)
                {
                    current +=
                        $"Стр.0 не должна быть меньше максимального значения из строк 1,2,3,4,5 - {minCountAll}" +
                        $"и не должна быть больше суммы строк 1,2,3,4,5 - {maxCountAll}." + Environment.NewLine;
                }

                if (countPhone < 0)
                {
                    current += "Стр.1 не должна быть меньше 0." + Environment.NewLine;
                }

                if (countSms < 0)
                {
                    current += "Стр.2 не должна быть меньше 0." + Environment.NewLine;
                }

                if (countPostal < 0)
                {
                    current += "Стр.4 не должна быть меньше 0." + Environment.NewLine;
                }

                if (countDisagree < 0)
                {
                    current += "Стр.6 не должна быть меньше 0." + Environment.NewLine;
                }

                if (countAll < 0)
                {
                    current += "Стр.0 должна быть больше 0." + Environment.NewLine;
                }

                if (countMessengers < countViber)
                {
                    current += "Стр.3 должна быть больше или равна стр.3-1." + Environment.NewLine;
                }

                if (countE < countEmail && countE < countMobileApp)
                {
                    current += "Стр.5 должна быть больше или равна любой из стр.5-1 или 5-2." + Environment.NewLine;
                }

                if (current.Length > 0)
                {
                    message += $"{data.Theme}. " + Environment.NewLine + current;
                }
            }

            foreach (var data in Report.ReportDataList.Where(x => x.Theme.StartsWith("Тема")))
            {
                if (data.Data == null)
                {
                    continue;
                }

                decimal sumPersFirst = 0;
                decimal sumPersRepeat = 0;

                string current = "";
                foreach (var iizl in data.Data)
                {
                    if (iizl.CountPersFirst == 0 && iizl.CountPersRepeat == 0)
                    {
                        continue;
                    }

                    if (iizl.CountMessages == 0)
                    {
                        current += $"Код {iizl.Code}. Необходимо заполнить гр.3" + Environment.NewLine;
                    }

                    if (iizl.CountMessages < iizl.CountPersFirst + iizl.CountPersRepeat)
                    {
                        current +=
                            $"Код {iizl.Code}. Количество сообщений не должно быть меньше суммы гр.1 и гр.2" +
                            Environment.NewLine;
                    }

                    if (iizl.CountMessages > 0 && iizl.TotalCost == 0)
                    {
                        current += $"Код {iizl.Code}. Необходимо указать сумму" + Environment.NewLine;
                    }

                    if (iizl.CountMessages > iizl.TotalCost * 100)
                    {
                        current += $"Код {iizl.Code}. Минимально возможная сумма - это 1 копейка * гр.3"
                                   + Environment.NewLine;
                    }

                    if (string.IsNullOrWhiteSpace(iizl.AccountingDocument))
                    {
                        current += $"Код {iizl.Code}. Необходимо заполнить гр.5" + Environment.NewLine;
                    }

                    if (iizl.CountMessages > 0 && iizl.CountPersFirst == 0 && iizl.CountPersRepeat == 0)
                    {
                        current += $"Код {iizl.Code}. Гр.3 должна быть равно 0, если гр.1 = 0 и гр.2 = 0" +
                                   Environment.NewLine;
                    }

                    sumPersFirst += iizl.CountPersFirst;
                    sumPersRepeat += iizl.CountPersRepeat;
                }

                if (sumPersFirst > 0 && data.TotalPersFirst == 0)
                {
                    current += "В строке 'Итого ЗЛ' гр.1 должна быть больше 0" + Environment.NewLine;
                }

                if (sumPersRepeat > 0 && data.TotalPersRepeat == 0)
                {
                    current += "В строке 'Итого ЗЛ' гр.2 должна быть больше 0" + Environment.NewLine;
                }

                if (sumPersFirst < data.TotalPersFirst)
                {
                    current += "В строке 'Итого ЗЛ' гр.1 должна быть меньше или равна сумме всех ячеек в столбце" +
                               Environment.NewLine;
                }

                if (sumPersRepeat < data.TotalPersRepeat)
                {
                    current += "В строке 'Итого ЗЛ' гр.2 должна быть меньше или равна сумме всех ячеек в столбце" +
                               Environment.NewLine;
                }

                if (current.Length > 0)
                {
                    message += $"{data.Theme}. " + Environment.NewLine + current;
                }
            }

            if (message.Length > 0)
            {
                message = "Форма ИИЗЛ. " + Environment.NewLine + message;
            }

            return message;
        }

        private void FillDgwInfAgree(DataGridView dgvReport, string form)
        {
            var reportIizlDto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form) ?? new ReportIizlDto();
            var data = reportIizlDto.Data;

            if (data.Length <= 0)
            {
                return;
            }

            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                string code = row.Cells[1].Value.ToString();
                var countPersFirst = data.Single(x => x.Code == code).CountPersFirst;
                row.Cells[2].Value = countPersFirst;
            }
        }

        private void FillDgwThemes(DataGridView dgvReport, string form)
        {
            

            var reportIizlDto = Report.ReportDataList.SingleOrDefault(x => x.Theme.ToLower() == form.ToLower());

            if (reportIizlDto == null)
            {
                return;
            }

            if (reportIizlDto.Data == null || reportIizlDto.Data.Length == 0)
            {
                return;
            }

            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                var code = row.Cells[1].Value.ToString();
                if (!code.StartsWith("И-"))
                {
                    var data = reportIizlDto.Data.Single(x => x.Code == code);
                    if (data != null)
                    {
                        row.Cells[2].Value = data.CountPersFirst;
                        row.Cells[3].Value = data.CountPersRepeat;
                        row.Cells[4].Value = data.CountMessages;
                        row.Cells[5].Value = Math.Round(data.TotalCost, 2);
                        row.Cells[6].Value = data.AccountingDocument;
                    }
                }
                else
                {
                    row.Cells[2].Value = reportIizlDto.TotalPersFirst;
                    row.Cells[3].Value = reportIizlDto.TotalPersRepeat;
                }
            }
        }

        private void CreateDgvInfAgreeColumns(DataGridView dgvReport)
        {
            var column = new DataGridViewTextBoxColumn {
                HeaderText = @"Наименование",
                Width = 450,
                DataPropertyName = "Naim",
                Name = "Naim",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle {BackColor = Color.Azure}
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn {
                HeaderText = @"Номер строки",
                Width = 100,
                DataPropertyName = "Num",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Name = "Num",
                DefaultCellStyle = new DataGridViewCellStyle {BackColor = Color.Azure}
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn {
                HeaderText = @"Количество",
                Width = 100,
                DataPropertyName = "Code",
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Name = "Code"
            };
            dgvReport.Columns.Add(column);
        }

        private void CreateDgvThemeColumns(DataGridView dgvReport)
        {
            var column = new DataGridViewTextBoxColumn {
                HeaderText = @"Способы информирования",
                Width = 265,
                DataPropertyName = "Way",
                Name = "Way",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle {BackColor = Color.Azure}
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn {
                HeaderText = @"Код",
                Width = 50,
                DataPropertyName = "Code",
                Name = "Code",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle {BackColor = Color.Azure}
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn {
                HeaderText = @"Количество ЗЛ (первичное информирование по теме)" + Environment.NewLine + @"(гр.1)",
                Width = 120,
                DataPropertyName = "CountPeopleFirst",
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Name = "CountPeopleFirst"
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn {
                HeaderText = @"Количество ЗЛ (повторное информирование по теме)" + Environment.NewLine + @"(гр.2)",
                Width = 120,
                DataPropertyName = "CountPeopleRepeat",
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Name = "CountPeopleRepeat"
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn {
                HeaderText = @"Количество сообщений (первичное и повторное информирование по теме)" +
                             Environment.NewLine + @"(гр.3)",
                Width = 120,
                DataPropertyName = "CountMessages",
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Name = "CountMessages"
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn {
                HeaderText = @"Суммарные затраты(руб.)" + Environment.NewLine + @"(гр.4)",
                Width = 120,
                DataPropertyName = "TotalCost",
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Name = "TotalCost"
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn {
                HeaderText = @"Реквизиты учетного документа" + Environment.NewLine + @"(гр.5)",
                Width = 120,
                DataPropertyName = "AccountingDocument",
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Name = "AccountingDocument"
            };
            dgvReport.Columns.Add(column);
        }
    }
}