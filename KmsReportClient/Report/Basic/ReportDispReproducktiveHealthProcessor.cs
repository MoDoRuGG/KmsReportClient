using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using KmsReportClient.Excel.Creator.Base;
using KmsReportClient.External;
using KmsReportClient.Forms;
using KmsReportClient.Global;
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.XML;
using KmsReportClient.Support;
using NLog;

namespace KmsReportClient.Report.Basic
{
    class ReportDispReproducktiveHealthProcessor : AbstractReportProcessor<ReportDispReprodHealth>
    {
        private static readonly string[] _notSaveCells =
            {
                "4", "5",
            };

        private static readonly Dictionary<string, string[]> _sumRules = new Dictionary<string, string[]>
        {

            ["4"] = new[] { "4.1", "4.2", "4.3", "4.4", "4.5", "4.6" },
            ["5"] = new[] { "5.1", "5.2", "5.3", "5.4", "5.5", "5.6" },
            
        };

        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        Dictionary<string, DataGridViewRow> _rows;
        private readonly string[] _forms1 = { "Дисп РЗ" };


        private readonly string[][] _headers = {
            new[]
            { "Всего с начала года", "Всего за отчетный период" },
            
        };

        private readonly Dictionary<string, string> _headersMap = new Dictionary<string, string>
        {
            { "Дисп РЗ", "Информирование граждан репродуктивного возраста о возможности прохождения диспансеризации по оценке репродуктивного здоровья" },
        };

        public ReportDispReproducktiveHealthProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
                    base(inClient, dgv, cmb, txtb, page,
                        XmlFormTemplate.DispRepHealth.GetDescription(),
                        Log,
                        ReportGlobalConst.ReportDispRepHealth,
                        reportsDictionary)
        {
            InitReport();
        }

        public override void InitReport()
        {
            Report = new ReportDispReprodHealth { ReportDataList = new ReportDispReprodHealthDto[ThemesList.Count], IdType = IdReportType };

            int i = 0;
            foreach (var theme in ThemesList.Select(x => x.Key))
            {
                Report.ReportDataList[i++] = new ReportDispReprodHealthDto { Theme = theme };
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
                    reportType = ReportType.DispRepHeal
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response as ReportDispReprodHealth;
        }

        public override void MapForAutoFill(AbstractReport report)
        {
            if (report == null)
            {
                return;
            }
            var inReport = report as ReportDispReprodHealth;

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
            if (_forms1.Contains(form))
            {
                FillDgvForms1(Dgv, form);
            }

            Dgv.DefaultCellStyle.BackColor = Color.Azure;

            SetFormula();
        }

        public void SetFormula()
        {

            foreach (var row in _rows.Reverse())
            {
                var key = row.Key;

                // Обработка суммирования
                if (_sumRules.TryGetValue(key, out var childKeys))
                {
                    row.Value.Cells[2].Value = SumCells(childKeys, 2);
                    row.Value.Cells[3].Value = SumCells(childKeys, 3);
                    continue;
                }
            }
        }

        private decimal SumCells(string[] keys, int columnIndex)
        {
            return keys
                .Where(k => _rows.ContainsKey(k))
                .Sum(k => GlobalUtils.TryParseDecimal(_rows[k].Cells[columnIndex].Value));
        }

        private void SetStyle()
        {

            foreach (DataGridViewRow row in Dgv.Rows)
            {


                string rowNum = row.Cells[1].Value.ToString();
                if (_notSaveCells.Contains(rowNum))
                {
                    row.DefaultCellStyle.BackColor = Color.LightGray;

                    row.ReadOnly = false;
                    row.DefaultCellStyle.Font = new Font(Dgv.DefaultCellStyle.Font, FontStyle.Bold);
                }
                row.Cells[2].Style.BackColor = Color.LightGray;

            }

        }


        protected override void FillReport(string form)
        {
            if (form == null)
            {
                return;
            }
            if (_forms1.Contains(form))
            {
                FillThemesForms1(Dgv, form);
            }
        }

        public override bool IsVisibleBtnDownloadExcel() => false;
        public override bool IsVisibleBtnHandle() => false;

        public override bool IsVisibleBtnSummary() => false;

        public override string ValidReport()
        {
            string message = "";

            if (message.Length > 0)
            {
                message = "Диспансеризация репродуктивного здоровья. " + Environment.NewLine + message;
            }
            return message;
        }

        public override void ToExcel(string filename, string filialName)
        {
            var mm = YymmUtils.GetMonth(Report.Yymm.Substring(2, 2)) + " 20" + Report.Yymm.Substring(0, 2);
            var excel = new ExcelDispReprodHealthCreator(filename, ExcelForm.dispRepHealth, mm, filialName);
            excel.CreateReport(Report, null);
        }

        public override void SaveToDb()
        {
            SetFormula();
            var request = new SaveReportRequest
            {
                Body = new SaveReportRequestBody

                {
                    filialCode = CurrentUser.FilialCode,
                    idUser = CurrentUser.IdUser,
                    report = Report,
                    yymm = Report.Yymm,
                    reportType = ReportType.DispRepHeal
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportDispReprodHealth;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
            Report.DataSource = response.DataSource;

        }

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
                    reportType = ReportType.DispRepHeal
                }
            };
            var response = Client.SaveReportDataSourceExcel(request).Body.SaveReportDataSourceExcelResult as ReportDispReprodHealth;
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
                    reportType = ReportType.DispRepHeal
                }
            };
            var response = Client.SaveReportDataSourceHandle(request).Body.SaveReportDataSourceHandleResult as ReportDispReprodHealth;
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
                    reportType = ReportType.DispRepHeal
                }
            };
            var response = Client.CollectSummaryReport(request);
            Report = response.Body.CollectSummaryReportResult as ReportDispReprodHealth;
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
            }
            SetStyle();
            _rows = new Dictionary<string, DataGridViewRow>();
            foreach (DataGridViewRow row in Dgv.Rows)
            {
                _rows.Add(row.Cells[1].Value.ToString(), row);
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

        private void FillThemesForms1(DataGridView dgvReport, string form)
        {
            var reportDto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportDto == null)
            {
                return;
            }

            reportDto.Data = (from DataGridViewRow row in dgvReport.Rows
                                 let rowNum = row.Cells[1].Value.ToString().Trim()
                                 where !IsNotNeedFillRow(form, rowNum)
                                 select new ReportDispReprodHealthDataDto
                                 {
                                     Code = rowNum,
                                     ForPeriod = GlobalUtils.TryParseInt(row.Cells[3].Value),
                                     YearlySum = GlobalUtils.TryParseInt(row.Cells[2].Value)
                                 }).ToArray();
            if (reportDto.Data.Length > 0) { SetFormula(); }

        }

        private void FillDgvForms1(DataGridView dgvReport, string form)
        {
            var reportDto = Report.ReportDataList?.SingleOrDefault(x => x.Theme == form);
            if (reportDto?.Data == null || reportDto.Data.Length == 0)
                return;

            // Подготовка справочников
            var rows = ThemeTextData.Tables_fromxml
                .Where(x => x.TableName_fromxml == form)
                .SelectMany(x => x.Rows_fromxml)
                .ToList();

            var rowsLookup = rows.ToDictionary(r => r.RowNum_fromxml, r => r.Exclusion_fromxml);
            var dataDict = reportDto.Data.ToDictionary(d => d.Code);

            // Собираем все rowNum для пакетного запроса
            var allRowNums = dgvReport.Rows.Cast<DataGridViewRow>()
                .Select(r => r.Cells[1].Value?.ToString().Trim())
                .Where(rn => !string.IsNullOrEmpty(rn))
                .ToArray();

            // === Подготовка rowNumbers как ArrayOfString ===
            var rowNumbersList = new ArrayOfString();
            foreach (var rn in allRowNums)
            {
                rowNumbersList.Add(rn);
            }

            // === ОДИН пакетный вызов ===
            var yearDataArray = Client.GetDispReprodHealthYearDataBatch(
                new GetDispReprodHealthYearDataBatchRequest(
                    new GetDispReprodHealthYearDataBatchRequestBody
                    {
                        yymm = Report.Yymm,
                        theme = form,
                        fillial = FilialCode,
                        rowNumbers = rowNumbersList
                    })).Body.GetDispReprodHealthYearDataBatchResult;

            // === Преобразуем в словарь для быстрого поиска ===
            var yearDataDict = new Dictionary<string, ReportDispReprodHealthDataDto>();
            if (yearDataArray != null)
            {
                foreach (var item in yearDataArray)
                {
                    // Используем RowNum и Data из нового DTO
                    yearDataDict[item.RowNum] = item.Data ?? new ReportDispReprodHealthDataDto { ForPeriod = 0 };
                }
            }

            // === Заполняем DataGridView ===
            dgvReport.SuspendLayout();
            try
            {
                foreach (DataGridViewRow row in dgvReport.Rows)
                {
                    var rowNum = row.Cells[1].Value?.ToString().Trim();
                    if (string.IsNullOrEmpty(rowNum)) continue;

                    bool isExclusionsRow = rowsLookup.TryGetValue(rowNum, out var excl) && excl;

                    // Заполняем из основных данных отчёта
                    if (dataDict.TryGetValue(rowNum, out var data))
                    {
                        if (rowNum != "7.5")
                            row.Cells[3].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.ForPeriod);
                        else
                            row.Cells[2].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.YearlySum);
                    }

                    // Заполняем из годовых данных (пакетная загрузка)
                    if (yearDataDict.TryGetValue(rowNum, out var yearData))
                    {
                        if (rowNum != "7.5")
                        {
                            // Убедитесь, что ячейка принимает decimal или строку
                            row.Cells[2].Value = yearData.ForPeriod;
                        }
                    }
                }
            }
            finally
            {
                dgvReport.ResumeLayout();
            }

            // Пересчитываем формулы один раз после заполнения
            SetFormula();
        }
    }
}
