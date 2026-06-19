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
    class ReportZpzT1Processor : AbstractReportProcessor<ReportZpz2025>
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();


        private readonly string[] _forms1 = { "Таблица 1" };

        private readonly string[][] _headers = {
            new[]
            { "Устные", "Письменные","По поручениям" } //1
        };

        private readonly Dictionary<string, string> _headersMap = new Dictionary<string, string>
        {
            { "Таблица 1", "Виды обращений" },
        };

        public ReportZpzT1Processor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
            base(inClient, dgv, cmb, txtb, page,
                XmlFormTemplate.ZpzT1.GetDescription(),
                Log,
                ReportGlobalConst.ReportZpzT1,
                reportsDictionary)
        {
            InitReport();
        }

        public override void InitReport()
        {
            Report = new ReportZpz2025 { ReportDataList = new ReportZpz2025Dto[ThemesList.Count], IdType = IdReportType };

            int i = 0;
            foreach (var theme in ThemesList.Select(x => x.Key))
            {
                Report.ReportDataList[i++] = new ReportZpz2025Dto { Theme = theme };
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
                    reportType = ReportType.ZpzT1
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response as ReportZpz2025;
        }

        public override void MapForAutoFill(AbstractReport report)
        {
            if (report == null)
            {
                return;
            }
            var inReport = report as ReportZpz2025;

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
            if (Report.DataSource != DataSource.Handle)
            {
                Dgv.DefaultCellStyle.BackColor = Color.LightGray;
            }
            else
            {
                Dgv.DefaultCellStyle.BackColor = Color.Azure;
            }
            SetTotalColumn();
        }

        protected override void FillReport(string form)
        {
            if (form == null)
            {
                return;
            }
            if (_forms1.Contains(form))
            {
                FillThemesForms3(Dgv, form);
            }
        }

        public override bool IsVisibleBtnDownloadExcel() => (CurrentUser.IsMain || Report.Status == ReportStatus.Done || Report.Status == ReportStatus.Submit) ? false : true;
        public override bool IsVisibleBtnHandle() => (CurrentUser.IsMain || Report.Status == ReportStatus.Done || Report.Status == ReportStatus.Submit) ? false : true;
        public override bool IsVisibleBtnSummary() => false;

        public override string ValidReport()
        {
            string message = "";
            if (FilialCode != "RU-PER")
            {
                // Получаем данные для "Таблица 1"
                var table1Data = Report.ReportDataList?.FirstOrDefault(x => x.Theme == "Таблица 1")?.Data;
                if (table1Data == null || table1Data.Length == 0)
                {
                    return message;
                }

                // Родительская строка 3.1.6
                var parentRow = table1Data.FirstOrDefault(x => x.Code == "3.1.6");
                if (parentRow == null)
                {
                    return message;
                }

                // Дочерние строки: 3.1.6.1 – 3.1.6.13
                // Фильтруем по префиксу "3.1.6." и проверяем, что код имеет формат "3.1.6.X"
                var childRows = table1Data.Where(x =>
                    x.Code.StartsWith("3.1.6.") &&
                    x.Code.Length > 6 &&
                    x.Code.Split('.').Length == 4).ToArray();

                // Суммируем значения по каждому столбцу
                // Примечание: свойства имеют тип decimal (не nullable), поэтому ?? не нужен
                var sumCountSmo = childRows.Sum(x => x.CountSmo);
                var sumCountSmoAnother = childRows.Sum(x => x.CountSmoAnother);
                var sumCountAssignment = childRows.Sum(x => x.CountAssignment);

                // Сравниваем с родительской строкой (с допуском для decimal)
                const decimal epsilon = 0.01m;

                if (Math.Abs(parentRow.CountSmo - sumCountSmo) > epsilon)
                {
                    message += $"гр.4 (Устные): значение стр.3.1.6 ({parentRow.CountSmo}) не равно сумме строк 3.1.6.1–3.1.6.13 ({sumCountSmo})\r\n";
                }

                if (Math.Abs(parentRow.CountSmoAnother - sumCountSmoAnother) > epsilon)
                {
                    message += $"гр.5 (Письменные): значение стр.3.1.6 ({parentRow.CountSmoAnother}) не равно сумме строк 3.1.6.1–3.1.6.13 ({sumCountSmoAnother})\r\n";
                }

                if (Math.Abs(parentRow.CountAssignment - sumCountAssignment) > epsilon)
                {
                    message += $"гр.6 (По поручениям): значение стр.3.1.6 ({parentRow.CountAssignment}) не равно сумме строк 3.1.6.1–3.1.6.13 ({sumCountAssignment})\r\n";
                }

                // Формируем итоговое сообщение
                if (!string.IsNullOrEmpty(message))
                {
                    message = "Таблица 1. \r\n" + message;
                }
            }
            return message;
            
        }

        public override void ToExcel(string filename, string filialName)
        {
            var mm = YymmUtils.GetMonth(Report.Yymm.Substring(2, 2)) + " 20" + Report.Yymm.Substring(0, 2);
            var excel = new ExcelZpz2025Creator(filename, ExcelForm.Zpz2025, mm, filialName);
            excel.CreateReport(Report, null);
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
                    reportType = ReportType.ZpzT1
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportZpz2025;
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
                    reportType = ReportType.ZpzT1
                }
            };
            var response = Client.SaveReportDataSourceExcel(request).Body.SaveReportDataSourceExcelResult as ReportZpz2025;
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
                    reportType = ReportType.ZpzT1
                }
            };
            var response = Client.SaveReportDataSourceHandle(request).Body.SaveReportDataSourceHandleResult as ReportZpz2025;
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
                    reportType = ReportType.ZpzT1
                }
            };
            var response = Client.CollectSummaryReport(request);
            Report = response.Body.CollectSummaryReportResult as ReportZpz2025;
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

            // ⚠️ КРИТИЧЕСКИ ВАЖНО: включаем перенос текста во ВСЕХ ячейках
            dgvReport.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvReport.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgvReport.RowTemplate.Height = 0; // разрешаем автоматическую высоту

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

        private void FillThemesForms3(DataGridView dgvReport, string form)
        {
            var reportZpz2025Dto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportZpz2025Dto == null)
            {
                return;
            }

            reportZpz2025Dto.Data = (from DataGridViewRow row in dgvReport.Rows
                                     let rowNum = row.Cells[1].Value.ToString().Trim()
                                     where !IsNotNeedFillRow(form, rowNum)
                                     select new ReportZpz2025DataDto
                                     {
                                         Code = rowNum,
                                         CountSmo = GlobalUtils.TryParseDecimal(row.Cells[2].Value),
                                         CountSmoAnother = GlobalUtils.TryParseDecimal(row.Cells[3].Value),
                                         CountAssignment = GlobalUtils.TryParseDecimal(row.Cells[4].Value)
                                     }).ToArray();
        }


        private void FillDgvForms1(DataGridView dgvReport, string form)
        {
            var reportZpz2025Dto = Report.ReportDataList?.SingleOrDefault(x => x.Theme == form);
            if (reportZpz2025Dto?.Data == null || reportZpz2025Dto.Data.Length == 0)
            {
                return;
            }

            var rows = ThemeTextData.Tables_fromxml.Where(x => x.TableName_fromxml == form).SelectMany(x => x.Rows_fromxml).ToList();
            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                var rowNum = row.Cells[1].Value.ToString().Trim();
                bool exclusionsRow = rows.Single(x => x.RowNum_fromxml == rowNum).Exclusion_fromxml;

                var data = reportZpz2025Dto.Data.SingleOrDefault(x => x.Code == rowNum);
                if (data != null)
                {
                    row.Cells[2].Value = exclusionsRow ? "x" : data.CountSmo.ToString().Replace(",00", "");
                    row.Cells[3].Value = exclusionsRow ? "x" : data.CountSmoAnother.ToString().Replace(",00", "");
                    row.Cells[4].Value = exclusionsRow ? "x" : data.CountAssignment.ToString().Replace(",00", "");
                }
            }
        }
    }
}
