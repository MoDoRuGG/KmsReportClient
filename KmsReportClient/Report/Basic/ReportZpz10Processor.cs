using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.ServiceModel.Channels;
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
    class ReportZpz10Processor : AbstractReportProcessor<ReportZpz>
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        private readonly string[] _forms1 = { "Таблица 10" };


        private readonly string[][] _headers = {
            new[]
            { "Всего" }, //10
        };

        private readonly Dictionary<string, string> _headersMap = new Dictionary<string, string>
        {
            { "Таблица 10", "Численность проинформированных застрахованных лиц" },
        };

        public ReportZpz10Processor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
            base(inClient, dgv, cmb, txtb, page,
                XmlFormTemplate.Zpz10.GetDescription(),
                Log,
                ReportGlobalConst.ReportZpz10,
                reportsDictionary)
        {
            InitReport();
        }

        public override void InitReport()
        {
            Report = new ReportZpz { ReportDataList = new ReportZpzDto[ThemesList.Count], IdType = IdReportType };

            int i = 0;
            foreach (var theme in ThemesList.Select(x => x.Key))
            {
                Report.ReportDataList[i++] = new ReportZpzDto { Theme = theme };
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
                    reportType = ReportType.Zpz10
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response as ReportZpz;
        }

        public override void MapForAutoFill(AbstractReport report)
        {
            if (report == null)
            {
                return;
            }
            var inReport = report as ReportZpz;

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
                FillDgwForms1(Dgv, form);
            }


            if (Report.DataSource != DataSource.Handle)
            {
                Dgv.DefaultCellStyle.BackColor = Color.LightGray;
            }
            else
            { Dgv.DefaultCellStyle.BackColor = Color.Azure; }
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
                FillThemesForms1(Dgv, form);
            }
        }

        public override string ValidReport()
        {
            string message = "";
            if (message.Length > 0)
            {
                message = "Форма ЗПЗ. " + Environment.NewLine + message;
            }
            return message;
        }

        public override bool IsVisibleBtnDownloadExcel() => true;

        public override bool IsVisibleBtnHandle() => true;

        public override void ToExcel(string filename, string filialName)
        {
            var mm = YymmUtils.GetMonth(Report.Yymm.Substring(2, 2)) + " 20" + Report.Yymm.Substring(0, 2);
            var excel = new ExcelZpz10Creator(filename, ExcelForm.Zpz10, mm, filialName);
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
                    reportType = ReportType.Zpz10
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportZpz;
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
                    reportType = ReportType.Zpz10
                }
            };
            var response = Client.SaveReportDataSourceExcel(request).Body.SaveReportDataSourceExcelResult as ReportZpz;
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
                    reportType = ReportType.Zpz10
                }
            };
            var response = Client.SaveReportDataSourceHandle(request).Body.SaveReportDataSourceHandleResult as ReportZpz;
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
                    reportType = ReportType.Zpz10
                }
            };
            var response = Client.CollectSummaryReport(request);
            Report = response.Body.CollectSummaryReportResult as ReportZpz;
            Report.IdType = IdReportType;
            Report.Yymm = yymmEnd;
        }

        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            var formsList = ThemesList.Select(x => x.Key).OrderBy(x => x).ToList();
            var index = formsList.IndexOf(form);
            var currentHeaders = _headers[index];
            CreateDgvColumnsForTheme(Dgv, 400, _headersMap[form], currentHeaders);

            int countRows = ThemeTextData.tables.Single(x => x.Name == form).RowsCount;
            foreach (var row in table)
            {
                var dgvRow = new DataGridViewRow();
                var cellName = new DataGridViewTextBoxCell
                {
                    Value = row.Name
                };
                var cellNum = new DataGridViewTextBoxCell
                {
                    Value = row.Num
                };
                dgvRow.Cells.Add(cellName);
                dgvRow.Cells.Add(cellNum);

                var exclusionCells = row.ExclusionCells?.Split(',');
                for (int i = 2; i < countRows; i++)
                {
                    bool isNeedExcludeSum = exclusionCells?.Contains(i.ToString()) ?? false;
                    var cell = new DataGridViewTextBoxCell
                    {
                        Value = row.Exclusion || isNeedExcludeSum ? "X" : "0"
                    };
                    dgvRow.Cells.Add(cell);

                    if (isNeedExcludeSum)
                    {
                        cell.ReadOnly = true;
                        cell.Style.BackColor = Color.DarkGray;
                    }
                }
                int rowIndex = Dgv.Rows.Add(dgvRow);
                if (row.Exclusion)
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

        private void FillThemesForms1(DataGridView dgvReport, string form)
        {
            var reportZpz10Dto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportZpz10Dto == null)
            {
                return;
            }

            reportZpz10Dto.Data = (from DataGridViewRow row in dgvReport.Rows
                                let rowNum = row.Cells[1].Value.ToString().Trim()
                                where !IsNotNeedFillRow(form, rowNum)
                                select new ReportZpzDataDto
                                {
                                    Code = rowNum,
                                    CountSmo = GlobalUtils.TryParseDecimal(row.Cells[2].Value),
                                    CountSmoAnother = GlobalUtils.TryParseDecimal(row.Cells[3].Value)
                                }).ToArray();
        }


        private void FillDgwForms1(DataGridView dgvReport, string form)
        {
            var reportZpz10Dto = Report.ReportDataList?.Single(x => x.Theme == form);
            if (reportZpz10Dto?.Data == null || reportZpz10Dto.Data.Length == 0)
            {
                return;
            }

            var rows = ThemeTextData.tables.Where(x => x.Name == form).SelectMany(x => x.Rows).ToList();
            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                var rowNum = row.Cells[1].Value.ToString().Trim();
                bool isExclusionsRow = rows.Single(x => x.Num == rowNum).Exclusion;

                var data = reportZpz10Dto.Data.SingleOrDefault(x => x.Code == rowNum);
                if (data != null)
                {
                    row.Cells[2].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.CountSmo);
                    row.Cells[3].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.CountSmoAnother);
                }
            }
        }
    }
}
