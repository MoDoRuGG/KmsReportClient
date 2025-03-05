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
    class ReportZpz10_2025Processor : AbstractReportProcessor<ReportZpz2025>
    {
        string[] _notSaveCells = new string[] { "1",
                                                "2",
                                                "3",
                                                "4", "4.1", "4.2", "4.3", "4.4", "4.5", "4.6",
                                                "5", "5.1", "5.2", "5.3", "5.4", "5.5", "5.6",
                                                "6",
                                                "7",
                                                "8"
                                                };

        string[] _notStyleCells = new string[] {"7.5",
                                                "8.1", "8.2", "8.3", "8.4", "8.5", "8.6"
                                                };

        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        Dictionary<string, DataGridViewRow> _rows;
        private readonly string[] _forms1 = { "Таблица 10", };
        private readonly string[] _forms2 = { "Сведения СП" };

        private readonly string[][] _headers = {
            new[]
            { "Штатные работники", "Привлекаемые по договору" }, //сведения СП
            new[]
            { "С начала года", "За отчетный период" }, //10
            
        };

        private readonly Dictionary<string, string> _headersMap = new Dictionary<string, string>
        {
            { "Таблица 10", "Численность проинформированных застрахованных лиц" },
            { "Сведения СП", "Сведения по страховым представителям" }
        };

        public ReportZpz10_2025Processor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
                    base(inClient, dgv, cmb, txtb, page,
                        XmlFormTemplate.Zpz10_2025.GetDescription(),
                        Log,
                        ReportGlobalConst.ReportZpz10_2025,
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
                    reportType = ReportType.Zpz10_2025
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
            var waitingForm = new WaitingForm();
            waitingForm.Show();
            Application.DoEvents();
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
            
            if (_forms2.Contains(form))
            {
                FillDgvForms2(Dgv, form);
                SetTotalColumn();
            }

            waitingForm.Close();
        }

        public void SetFormula()
        {
            if (GetCurrentTheme() != "Сведения СП")
            {
                foreach (var row in _rows.Reverse())
                {
                    if (row.Key == "1")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "1.1" || x.Key == "1.2" || x.Key == "1.3" || x.Key == "1.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "1.1" || x.Key == "1.2" || x.Key == "1.3" || x.Key == "1.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }

                    if (row.Key == "2")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.1" || x.Key == "2.2" || x.Key == "2.3" || x.Key == "2.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.1" || x.Key == "2.2" || x.Key == "2.3" || x.Key == "2.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }

                    if (row.Key == "3")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "3.1" || x.Key == "3.2" || x.Key == "3.3" || x.Key == "3.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "3.1" || x.Key == "3.2" || x.Key == "3.3" || x.Key == "3.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }

                    if (row.Key == "4")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "4.1" || x.Key == "4.2" || x.Key == "4.3" || x.Key == "4.4" || x.Key == "4.5" || x.Key == "4.6").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "4.1" || x.Key == "4.2" || x.Key == "4.3" || x.Key == "4.4" || x.Key == "4.5" || x.Key == "4.6").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }

                    if (row.Key == "4.1")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "4.1.1" || x.Key == "4.1.2" || x.Key == "4.1.3" || x.Key == "4.1.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "4.1.1" || x.Key == "4.1.2" || x.Key == "4.1.3" || x.Key == "4.1.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }
                    if (row.Key == "4.2")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "4.2.1" || x.Key == "4.2.2" || x.Key == "4.2.3" || x.Key == "4.2.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "4.2.1" || x.Key == "4.2.2" || x.Key == "4.2.3" || x.Key == "4.2.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }
                    if (row.Key == "4.3")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "4.3.1" || x.Key == "4.3.2" || x.Key == "4.3.3" || x.Key == "4.3.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "4.3.1" || x.Key == "4.3.2" || x.Key == "4.3.3" || x.Key == "4.3.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }
                    if (row.Key == "4.4")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "4.4.1" || x.Key == "4.4.2" || x.Key == "4.4.3" || x.Key == "4.4.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "4.4.1" || x.Key == "4.4.2" || x.Key == "4.4.3" || x.Key == "4.4.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }
                    if (row.Key == "4.5")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "4.5.1" || x.Key == "4.5.2" || x.Key == "4.5.3" || x.Key == "4.5.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "4.5.1" || x.Key == "4.5.2" || x.Key == "4.5.3" || x.Key == "4.5.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }
                    if (row.Key == "4.6")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "4.6.1" || x.Key == "4.6.2" || x.Key == "4.6.3" || x.Key == "4.6.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "4.6.1" || x.Key == "4.6.2" || x.Key == "4.6.3" || x.Key == "4.6.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }
                    if (row.Key == "5")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "5.1" || x.Key == "5.2" || x.Key == "5.3" || x.Key == "5.4" || x.Key == "5.5" || x.Key == "5.6").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "5.1" || x.Key == "5.2" || x.Key == "5.3" || x.Key == "5.4" || x.Key == "5.5" || x.Key == "5.6").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }
                    if (row.Key == "5.1")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "5.1.1" || x.Key == "5.1.2" || x.Key == "5.1.3" || x.Key == "5.1.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "5.1.1" || x.Key == "5.1.2" || x.Key == "5.1.3" || x.Key == "5.1.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }
                    if (row.Key == "5.2")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "5.2.1" || x.Key == "5.2.2" || x.Key == "5.2.3" || x.Key == "5.2.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "5.2.1" || x.Key == "5.2.2" || x.Key == "5.2.3" || x.Key == "5.2.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }
                    if (row.Key == "5.3")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "5.3.1" || x.Key == "5.3.2" || x.Key == "5.3.3" || x.Key == "5.3.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "5.3.1" || x.Key == "5.3.2" || x.Key == "5.3.3" || x.Key == "5.3.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }
                    if (row.Key == "5.4")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "5.4.1" || x.Key == "5.4.2" || x.Key == "5.4.3" || x.Key == "5.4.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "5.4.1" || x.Key == "5.4.2" || x.Key == "5.4.3" || x.Key == "5.4.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }
                    if (row.Key == "5.5")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "5.5.1" || x.Key == "5.5.2" || x.Key == "5.5.3" || x.Key == "5.5.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "5.5.1" || x.Key == "5.5.2" || x.Key == "5.5.3" || x.Key == "5.5.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }
                    if (row.Key == "5.6")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "5.6.1" || x.Key == "5.6.2" || x.Key == "5.6.3" || x.Key == "5.6.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "5.6.1" || x.Key == "5.6.2" || x.Key == "5.6.3" || x.Key == "5.6.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }
                    if (row.Key == "6")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "6.1" || x.Key == "6.2" || x.Key == "6.3" || x.Key == "6.4" || x.Key == "6.5" || x.Key == "6.6" || x.Key == "6.7").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "6.1" || x.Key == "6.2" || x.Key == "6.3" || x.Key == "6.4" || x.Key == "6.5" || x.Key == "6.6" || x.Key == "6.7").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }
                    if (row.Key == "7")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "7.1" || x.Key == "7.2" || x.Key == "7.3" || x.Key == "7.4" || x.Key == "7.5" || x.Key == "7.6" || x.Key == "7.7" || x.Key == "7.8").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "7.1" || x.Key == "7.2" || x.Key == "7.3" || x.Key == "7.4" || x.Key == "7.5" || x.Key == "7.6" || x.Key == "7.7" || x.Key == "7.8").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }
                    if (row.Key == "8")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "8.1" || x.Key == "8.2" || x.Key == "8.3" || x.Key == "8.4" || x.Key == "8.5" || x.Key == "8.6").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "8.1" || x.Key == "8.2" || x.Key == "8.3" || x.Key == "8.4" || x.Key == "8.5" || x.Key == "8.6").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        continue;
                    }
                    if (row.Key == "7.5")
                    {
                        row.Value.Cells[3].Value = "X";
                        row.Value.Cells[3].ReadOnly = true;
                        row.Value.Cells[3].Style.BackColor = Color.LightGray;
                    }
                }
            }
        }

        private void SetStyle()
        {
            if (GetCurrentTheme() != "Сведения СП")
            {
                foreach (DataGridViewRow row in Dgv.Rows)
                {
                    if (_notStyleCells.Contains(row.Cells[1].Value.ToString()))
                    {
                        continue;
                    }
                    else
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
                    if (row.Cells[1].Value.ToString() == "7.5")
                    {
                        row.Cells[3].Style.BackColor = row.Cells[2].Style.BackColor = Color.Azure;
                        row.ReadOnly = false;
                    }
                }
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
            if (_forms2.Contains(form))
            {
                FillThemesForms2(Dgv, form);
            }
        }

        public override bool IsVisibleBtnDownloadExcel() => (CurrentUser.IsMain || Report.Status == ReportStatus.Done || Report.Status == ReportStatus.Submit) ? false : true;
        public override bool IsVisibleBtnHandle() => (CurrentUser.IsMain || Report.Status == ReportStatus.Done || Report.Status == ReportStatus.Submit) ? false : true;

        public override bool IsVisibleBtnSummary() => false;

        public override string ValidReport()
        {
            string message = "";

            if (message.Length > 0)
            {
                message = "Форма ЗПЗ 2025. " + Environment.NewLine + message;
            }
            return message;
        }

        public override void ToExcel(string filename, string filialName)
        {
            var mm = YymmUtils.GetMonth(Report.Yymm.Substring(2, 2)) + " 20" + Report.Yymm.Substring(0, 2);
            var excel = new ExcelZpz10_2025Creator(filename, ExcelForm.Zpz10_2025, mm, filialName);
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
                    reportType = ReportType.Zpz10_2025
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
                    reportType = ReportType.Zpz10_2025
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
                    reportType = ReportType.Zpz10_2025
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
                    reportType = ReportType.Zpz10_2025
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
            var reportZpzDto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportZpzDto == null)
            {
                return;
            }

            reportZpzDto.Data = (from DataGridViewRow row in dgvReport.Rows
                                 let rowNum = row.Cells[1].Value.ToString().Trim()
                                 where !IsNotNeedFillRow(form, rowNum)
                                 select new ReportZpz2025DataDto
                                 {
                                     Code = rowNum,
                                     CountSmo = GlobalUtils.TryParseDecimal(row.Cells[3].Value),
                                     CountSmoAnother = GlobalUtils.TryParseDecimal(row.Cells[2].Value)
                                 }).ToArray();
            if (reportZpzDto.Data.Length > 0) { SetFormula(); }

        }

        private void FillDgvForms1(DataGridView dgvReport, string form)
        {
            var reportZpzDto = Report.ReportDataList?.Single(x => x.Theme == form);
            if (reportZpzDto?.Data == null || reportZpzDto.Data.Length == 0)
            {
                return;
            }

            var rows = ThemeTextData.Tables_fromxml.Where(x => x.TableName_fromxml == form).SelectMany(x => x.Rows_fromxml).ToList();



            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                var rowNum = row.Cells[1].Value.ToString().Trim();
                bool isExclusionsRow = rows.Single(x => x.RowNum_fromxml == rowNum).Exclusion_fromxml;

                var data = reportZpzDto.Data.SingleOrDefault(x => x.Code == rowNum);
                if (data != null)
                {
                    if (rowNum != "7.5") { row.Cells[3].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.CountSmo); }
                    else { row.Cells[2].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.CountSmoAnother); }
                }


                var yearThemeData = Client.GetZpz10_2025YearData(new GetZpz10_2025YearDataRequest(new GetZpz10_2025YearDataRequestBody
                {
                    fillial = FilialCode,
                    theme = form,
                    yymm = Report.Yymm,
                    rowNum = rowNum
                })).Body.GetZpz10_2025YearDataResult;
                if (yearThemeData != null)
                {
                    if (rowNum != "7.5") { row.Cells[2].Value = yearThemeData.CountSmo; }
                }
            }
            SetFormula();
        }



        private void FillThemesForms2(DataGridView dgvReport, string form)
        {
            var reportZpzDto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportZpzDto == null)
            {
                return;
            }

            reportZpzDto.Data = (from DataGridViewRow row in dgvReport.Rows
                                 let rowNum = row.Cells[1].Value.ToString()
                                 select new ReportZpz2025DataDto
                                 {
                                     Code = rowNum,
                                     CountSmo = GlobalUtils.TryParseDecimal(row.Cells[2].Value),
                                     CountSmoAnother = GlobalUtils.TryParseDecimal(row.Cells[3].Value)
                                 }).ToArray();

        }

        private void FillDgvForms2(DataGridView dgvReport, string form)
        {
            var reportZpzDto = Report.ReportDataList?.Single(x => x.Theme == form);
            if (reportZpzDto?.Data == null || reportZpzDto.Data.Length == 0)
            {
                return;
            }

            var rows = ThemeTextData.Tables_fromxml.Where(x => x.TableName_fromxml == form).SelectMany(x => x.Rows_fromxml).ToList();
            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                var rowNum = row.Cells[1].Value.ToString().Trim();
                bool isExclusionsRow = rows.Single(x => x.RowNum_fromxml == rowNum).Exclusion_fromxml;

                var data = reportZpzDto.Data.SingleOrDefault(x => x.Code == rowNum);
                if (data != null)
                {
                    row.Cells[2].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.CountSmo);
                    row.Cells[3].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.CountSmoAnother);
                }
            }
        }
    }
}
