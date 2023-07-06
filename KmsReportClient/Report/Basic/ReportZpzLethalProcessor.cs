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
    public class ReportZpzLethalProcessor : AbstractReportProcessor<ReportZpz>
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        private readonly string[] _forms6 = { "Таблица 1Л", "Таблица 2Л" };

        private readonly string[][] _headers = 
        {
            new[]
            {
                "амбулаторно - поликлинической медицинской помощи",
                "стационарной медицинской помощи",
                "стационаро - замещающей медицинской помощи",
                "скорой медицинской помощи вне медицинской организации",
                "Всего"

            }, //1Л
            new[]
            {
                "амбулаторно - поликлинической медицинской помощи",
                "стационарной медицинской помощи",
                "стационаро - замещающей медицинской помощи",
                "скорой медицинской помощи вне медицинской организации",
                "Всего"

            }, //2Л
        };

        private readonly Dictionary<string, string> _headersMap = new Dictionary<string, string>
        {
            { "Таблица 1Л", "" },
            { "Таблица 2Л", "" }
            };

        public override void SaveReportDataSourceExcel()
        { }
        public override void SaveReportDataSourceHandle()
        { }

        public ReportZpzLethalProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
            base(inClient, dgv, cmb, txtb, page,
                XmlFormTemplate.ZpzLethal.GetDescription(),
                Log,
                ReportGlobalConst.ReportZpzLethal,
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
                    reportType = ReportType.ZpzLethal
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
            if (_forms6.Contains(form))
            {
                FillDgwForms6(Dgv, form);
                SetTotalColumn();
            }
        }

        protected override void FillReport(string form)
        {
            if (form == null)
            {
                return;
            }
            if (_forms6.Contains(form))
            {
                FillThemesForms6(Dgv, form);
            }
        }

        public override bool IsVisibleBtnDownloadExcel() => true;
        public override bool IsVisibleBtnHandle() => false;


        public override string ValidReport()
        { 
            var message = "";
            return message;
        }

        public override void ToExcel(string filename, string filialName)
        {
            var mm = YymmUtils.GetMonth(Report.Yymm.Substring(2, 2)) + " 20" + Report.Yymm.Substring(0, 2);
            var excel = new ExcelZpzQCreator(filename, ExcelForm.ZpzQ, mm, filialName);
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
                    reportType = ReportType.ZpzLethal
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportZpz;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
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
                    reportType = ReportType.ZpzLethal
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
            if (GetCurrentTheme() == "Таблица 1Л" || GetCurrentTheme() == "Таблица 2Л")
            {
                CreateDgvColumnsForTheme(Dgv, 400, _headersMap[form], currentHeaders);
            }
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
                    bool isNeedExludeSum = exclusionCells?.Contains(i.ToString()) ?? false;
                    var cell = new DataGridViewTextBoxCell
                    {
                        Value = row.Exclusion_fromxml || isNeedExludeSum ? "X" : "0"
                    };
                    dgvRow.Cells.Add(cell);

                    if (isNeedExludeSum)
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

        private void CreateDgvColumnsForTheme(DataGridView dgvReport,
                                              int widthFirstColumn,
                                              string mainHeader,
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

            Console.WriteLine("Кол-во столбцов="+dgvReport.Columns.Count);
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
                Width = 70,
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


        private void FillThemesForms6(DataGridView dgvReport, string form)
        {
            var reportZpzDto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportZpzDto != null)
            {
                reportZpzDto.Data = (from DataGridViewRow row in dgvReport.Rows
                                     let rowNum = row.Cells[1].Value.ToString().Trim()
                                     where !IsNotNeedFillRow(form, rowNum)
                                     select new ReportZpzDataDto
                                     {
                                         Code = rowNum,
                                         CountAmbulatory = GlobalUtils.TryParseDecimal(row.Cells[2].Value),
                                         CountStac = GlobalUtils.TryParseDecimal(row.Cells[3].Value),
                                         CountDs = GlobalUtils.TryParseDecimal(row.Cells[4].Value),
                                         CountOutOfSmoAnother = GlobalUtils.TryParseDecimal(row.Cells[5].Value),
                                         CountSmo = GlobalUtils.TryParseDecimal(row.Cells[6].Value),

                                     }).ToArray();
            }
        }

        private void FillDgwForms6(DataGridView dgvReport, string form)
        {
            var reportZpzDto = Report.ReportDataList.FirstOrDefault(x => x.Theme == form);
            if (reportZpzDto == null)
            {
                return;
            }
            if (reportZpzDto.Data == null || reportZpzDto.Data.Length == 0)
            {
                return;
            }

            var rows = ThemeTextData.Tables_fromxml.Where(x => x.TableName_fromxml == form).SelectMany(x => x.Rows_fromxml).ToList();

            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                var rowNum = row.Cells[1].Value.ToString().Trim();
                var data = reportZpzDto.Data.SingleOrDefault(x => x.Code == rowNum);
                var exclusionsCells = rows.Single(x => x.RowNum_fromxml == rowNum).ExclusionCells_fromxml?.Split(',');
                bool isExclusionsRow = rows.Single(x => x.RowNum_fromxml == rowNum).Exclusion_fromxml;
                if (data == null)
                {
                    continue;
                }

                row.Cells[2].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 2, data.CountAmbulatory);
                row.Cells[3].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 3, data.CountStac);
                row.Cells[4].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 4, data.CountDs);
                row.Cells[5].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 5, data.CountOutOfSmoAnother);
                row.Cells[6].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 6, data.CountSmo);
            }
        }

        public void SetFormula()
        {
        }
    }
}
