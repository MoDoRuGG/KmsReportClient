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
    class ReportZpz10Processor : AbstractReportProcessor<ReportZpz>
    {
        string[] _notSaveCells = new string[] { "1", "2", "2.1", "2.2", "2.3", "2.4", "2.5", "2.6", "3", "4" };
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        Dictionary<string, DataGridViewRow> _rows;
        private readonly string[] _forms1 = { "Таблица 10" };

        private readonly string[][] _headers = {
            new[]
            { "С начала года", "За отчетный период" }, //10
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


            Dgv.DefaultCellStyle.BackColor = Color.Azure;

            SetFormula();
            //SetTotalColumn();
        }

        public void SetFormula()
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
                    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.1" || x.Key == "2.2" || x.Key == "2.3" || x.Key == "2.4" || x.Key == "2.5" || x.Key == "2.6").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.1" || x.Key == "2.2" || x.Key == "2.3" || x.Key == "2.4" || x.Key == "2.5" || x.Key == "2.6").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                    continue;
                }


                if (row.Key == "2.1")
                {
                    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.1.1" || x.Key == "2.1.2" || x.Key == "2.1.3" || x.Key == "2.1.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.1.1" || x.Key == "2.1.2" || x.Key == "2.1.3" || x.Key == "2.1.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                    continue;

                }

                if (row.Key == "2.2")
                {
                    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.2.1" || x.Key == "2.2.2" || x.Key == "2.2.3" || x.Key == "2.2.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.2.1" || x.Key == "2.2.2" || x.Key == "2.2.3" || x.Key == "2.2.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                    continue;

                }

                if (row.Key == "2.3")
                {
                    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.3.1" || x.Key == "2.3.2" || x.Key == "2.3.3" || x.Key == "2.3.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.3.1" || x.Key == "2.3.2" || x.Key == "2.3.3" || x.Key == "2.3.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                    continue;

                }

                if (row.Key == "2.4")
                {
                    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.4.1" || x.Key == "2.4.2" || x.Key == "2.4.3" || x.Key == "2.4.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.4.1" || x.Key == "2.4.2" || x.Key == "2.4.3" || x.Key == "2.4.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                    continue;

                }

                if (row.Key == "2.5")
                {
                    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.5.1" || x.Key == "2.5.2" || x.Key == "2.5.3" || x.Key == "2.5.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value)); 
                    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.5.1" || x.Key == "2.5.2" || x.Key == "2.5.3" || x.Key == "2.5.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                    continue;

                }

                if (row.Key == "2.6")
                {
                    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.6.1" || x.Key == "2.6.2" || x.Key == "2.6.3" || x.Key == "2.6.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.6.1" || x.Key == "2.6.2" || x.Key == "2.6.3" || x.Key == "2.6.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                    continue;

                }

                if (row.Key == "3")
                {
                    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "3.1" || x.Key == "3.2" || x.Key == "3.3" || x.Key == "3.4" || x.Key == "3.5" || x.Key == "3.6").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "3.1" || x.Key == "3.2" || x.Key == "3.3" || x.Key == "3.4" || x.Key == "3.5" || x.Key == "3.6").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                    continue;
                }

                if (row.Key == "4")
                {
                    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "4.1" || x.Key == "4.2" || x.Key == "4.3" || x.Key == "4.4" || x.Key == "4.5" || x.Key == "4.6" || x.Key == "4.7" || x.Key == "4.8").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "4.1" || x.Key == "4.2" || x.Key == "4.3" || x.Key == "4.4" || x.Key == "4.5" || x.Key == "4.6" || x.Key == "4.7" || x.Key == "4.8").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                    continue;
                }
            }
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

        public override string ValidReport()
        {
            string message = "";

            if (message.Length > 0)
            {
                message = "Форма ЗПЗ. " + Environment.NewLine + message;
            }
            return message;
        }

        public override void ToExcel(string filename, string filialName)
        {
            var mm = YymmUtils.GetMonth(Report.Yymm.Substring(2, 2)) + " 20" + Report.Yymm.Substring(0, 2);
            var excel = new ExcelZpz10Creator(filename, ExcelForm.Zpz10, mm, filialName);
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
                                 select new ReportZpzDataDto
                                 {
                                     Code = rowNum,
                                     CountSmo = GlobalUtils.TryParseDecimal(row.Cells[3].Value),
                                     //CountSmoAnother = GlobalUtils.TryParseDecimal(row.Cells[3].Value)
                                 }).ToArray();
            SetFormula();
        }

        private void FillDgwForms1(DataGridView dgvReport, string form)
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
                    row.Cells[3].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.CountSmo);
                }


                var yearThemeData = Client.GetZpz10YearData(new GetZpz10YearDataRequest(new GetZpz10YearDataRequestBody
                {
                    fillial = FilialCode,
                    theme = form,
                    yymm = Report.Yymm,
                    rowNum = rowNum
                })).Body.GetZpz10YearDataResult;
                if (yearThemeData != null)
                {
                   row.Cells[2].Value = yearThemeData.CountSmo;

                }
            }
            SetFormula();
        }
    }
}
