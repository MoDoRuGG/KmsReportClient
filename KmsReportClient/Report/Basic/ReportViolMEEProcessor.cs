using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using KmsReportClient.Excel.Creator.Base;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Model;
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.XML;
using KmsReportClient.Support;
using NLog;

namespace KmsReportClient.Report.Basic
{
    class ReportViolMEEProcessor : AbstractReportProcessor<ReportViolations>
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        Dictionary<string, DataGridViewRow> _rows;
        private readonly string[] _forms1 = { "Нарушения МЭЭ" };

        private readonly string[][] _headers = {
            new[]
            { $"За отчетный период" }, //МЭЭ
        };

        private readonly Dictionary<string, string> _headersMap = new Dictionary<string, string>
        {
            { "Нарушения МЭЭ", "Перечень оснований для отказа в оплате медицинской помощи"},
        };

        public ReportViolMEEProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
            base(inClient, dgv, cmb, txtb, page,
                XmlFormTemplate.ViolMEE.GetDescription(),
                Log,
                ReportGlobalConst.ReportViolMEE,
                reportsDictionary)
        {
            InitReport();
        }

        public override void InitReport()
        {
            Report = new ReportViolations { ReportDataList = new ReportViolationsDto[ThemesList.Count], IdType = IdReportType };

            
            int i = 0;
            foreach (var theme in ThemesList.Select(x => x.Key))
            {
                Report.ReportDataList[i++] = new ReportViolationsDto { Theme = theme };
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
                    reportType = ReportType.ViolMEE
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response as ReportViolations;
        }

        public override void MapForAutoFill(AbstractReport report)
        {
            if (report == null)
            {
                return;
            }
            var inReport = report as ReportViolations;

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
                FillDgvForms(Dgv, form);
            }


            //Dgv.DefaultCellStyle.BackColor = Color.Azure;

            //SetFormula();
            //SetTotalColumn();
        }

        public void SetFormula()
        {
            foreach (var row in _rows.Reverse())
            {
                //if (row.Key == "1")
                //{
                //    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "1.1" || x.Key == "1.2" || x.Key == "1.3" || x.Key == "1.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                //    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "1.1" || x.Key == "1.2" || x.Key == "1.3" || x.Key == "1.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                //    continue;
                //}

                //if (row.Key == "2")
                //{
                //    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.1" || x.Key == "2.2" || x.Key == "2.3" || x.Key == "2.4" || x.Key == "2.5" || x.Key == "2.6").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                //    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.1" || x.Key == "2.2" || x.Key == "2.3" || x.Key == "2.4" || x.Key == "2.5" || x.Key == "2.6").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                //    continue;
                //}


                //if (row.Key == "2.1")
                //{
                //    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.1.1" || x.Key == "2.1.2" || x.Key == "2.1.3" || x.Key == "2.1.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                //    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.1.1" || x.Key == "2.1.2" || x.Key == "2.1.3" || x.Key == "2.1.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                //    continue;

                //}

                //if (row.Key == "2.2")
                //{
                //    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.2.1" || x.Key == "2.2.2" || x.Key == "2.2.3" || x.Key == "2.2.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                //    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.2.1" || x.Key == "2.2.2" || x.Key == "2.2.3" || x.Key == "2.2.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                //    continue;

                //}

                //if (row.Key == "2.3")
                //{
                //    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.3.1" || x.Key == "2.3.2" || x.Key == "2.3.3" || x.Key == "2.3.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                //    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.3.1" || x.Key == "2.3.2" || x.Key == "2.3.3" || x.Key == "2.3.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                //    continue;

                //}

                //if (row.Key == "2.4")
                //{
                //    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.4.1" || x.Key == "2.4.2" || x.Key == "2.4.3" || x.Key == "2.4.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                //    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.4.1" || x.Key == "2.4.2" || x.Key == "2.4.3" || x.Key == "2.4.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                //    continue;

                //}

                //if (row.Key == "2.5")
                //{
                //    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.5.1" || x.Key == "2.5.2" || x.Key == "2.5.3" || x.Key == "2.5.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value)); 
                //    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.5.1" || x.Key == "2.5.2" || x.Key == "2.5.3" || x.Key == "2.5.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                //    continue;

                //}

                //if (row.Key == "2.6")
                //{
                //    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.6.1" || x.Key == "2.6.2" || x.Key == "2.6.3" || x.Key == "2.6.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                //    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.6.1" || x.Key == "2.6.2" || x.Key == "2.6.3" || x.Key == "2.6.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                //    continue;

                //}

                //if (row.Key == "3")
                //{
                //    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "3.1" || x.Key == "3.2" || x.Key == "3.3" || x.Key == "3.4" || x.Key == "3.5" || x.Key == "3.6").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                //    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "3.1" || x.Key == "3.2" || x.Key == "3.3" || x.Key == "3.4" || x.Key == "3.5" || x.Key == "3.6").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                //    continue;
                //}

                //if (row.Key == "4")
                //{
                //    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "4.1" || x.Key == "4.2" || x.Key == "4.3" || x.Key == "4.4" || x.Key == "4.5" || x.Key == "4.6" || x.Key == "4.7" || x.Key == "4.8").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                //    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "4.1" || x.Key == "4.2" || x.Key == "4.3" || x.Key == "4.4" || x.Key == "4.5" || x.Key == "4.6" || x.Key == "4.7" || x.Key == "4.8").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                //    continue;
                //}
            }
        }

        private void SetStyle()
        {
            foreach (DataGridViewRow row in Dgv.Rows)
            {
                string rowNum = row.Cells[1].Value.ToString();
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
                FillThemesForms(Dgv, form);
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
                message = "Форма Нарушения МЭЭ. " + Environment.NewLine + message;
            }
            return message;
        }

        public override void ToExcel(string filename, string filialName)
        {
            var mm = YymmUtils.GetMonth(Report.Yymm.Substring(2, 2)) + " 20" + Report.Yymm.Substring(0, 2);
            var excel = new ExcelViolMEECreator(filename, ExcelForm.ViolMEE, mm, filialName);
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
                    reportType = ReportType.ViolMEE
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportViolations;
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
                    reportType = ReportType.ViolMEE
                }
            };
            var response = Client.SaveReportDataSourceExcel(request).Body.SaveReportDataSourceExcelResult as ReportViolations;
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
                    reportType = ReportType.ViolMEE
                }
            };
            var response = Client.SaveReportDataSourceHandle(request).Body.SaveReportDataSourceHandleResult as ReportViolations;
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
                    reportType = ReportType.ViolMEE
                }
            };
            var response = Client.CollectSummaryReport(request);
            Report = response.Body.CollectSummaryReportResult as ReportViolations;
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
                    Width = 120,
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
                HeaderText = "Код нарушения",
                Width = 80,
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
            var reportViolMEEDto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportViolMEEDto == null)
            {
                return;
            }

            reportViolMEEDto.Data = (from DataGridViewRow row in dgvReport.Rows
                                 let rowNum = row.Cells[1].Value.ToString().Trim()
                                 where !IsNotNeedFillRow(form, rowNum)
                                 select new ReportViolationsDataDto
                                 {
                                     Code = rowNum,
                                     Count = GlobalUtils.TryParseDecimal(row.Cells[2].Value),
                                    
                                 }).ToArray();
            SetFormula();
        }

        private void FillDgvForms(DataGridView dgvReport, string form)
        {
            var reportViolMEEDto = Report.ReportDataList?.Single(x => x.Theme == form);
            if (reportViolMEEDto?.Data == null || reportViolMEEDto.Data.Length == 0)
            {
                return;
            }

            var rows = ThemeTextData.Tables_fromxml.Where(x => x.TableName_fromxml == form).SelectMany(x => x.Rows_fromxml).ToList();



            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                var rowNum = row.Cells[1].Value.ToString().Trim();
                bool isExclusionsRow = rows.Single(x => x.RowNum_fromxml == rowNum).Exclusion_fromxml;

                var data = reportViolMEEDto.Data.SingleOrDefault(x => x.Code == rowNum);
                if (data != null)
                {
                    row.Cells[2].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.Count);
                }



            }
            SetFormula();
        }
    }
}
