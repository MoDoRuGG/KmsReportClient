using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KmsReportClient.DgvHeaderGenerator;
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
    public class ReportEffectivenessProcessor : AbstractReportProcessor<ReportEffectiveness>
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        private readonly string[][] _headers =  // заголовки для стоблцов с данными
        {
            new[]
            {
            "ФИО врача-эксперта",
            "Занятость ставки",
            "Cпециальность для эксперта качества",
            "Вид  проводимой экспертизы (МЭЭ, ЭКМП)",
            "МЭЭ;План по количеству",
            "МЭЭ;Факт",
            "МЭЭ;% выполнения",
            "МЭЭ;План по доходам",
            "МЭЭ;Факт",
            "МЭЭ;% выполнения",
            "ЭКМП;План по количеству",
            "ЭКМП;Факт",
            "ЭКМП;% выполнения",
            "ЭКМП;План по доходам",
            "ЭКМП;Факт",
            "ЭКМП;% выполнения",
            },
        };

        private readonly Dictionary<string, string> _headersMap = new Dictionary<string, string>
        {
            { "Эффективность","№"},  // заголовок для 1 колонки
        };

        public ReportEffectivenessProcessor
        (
            EndpointSoap inClient,
            List<KmsReportDictionary> reportsDictionary,
            DataGridView dgv,
            ComboBox cmb,
            TextBox txtb,
            TabPage page
        ) :
        base
        (
            inClient, dgv, cmb, txtb, page,
            XmlFormTemplate.Effectiveness.GetDescription(),
            Log,
            ReportGlobalConst.ReportEffectiveness,
            reportsDictionary
        )
        {
            InitReport();
        }

        public override void InitReport()
        {
            Report = new ReportEffectiveness { ReportDataList = new ReportEffectivenessDto[ThemesList.Count], IdType = IdReportType };

            int i = 0;
            foreach (var theme in ThemesList.Select(x => x.Key))
            {
                Report.ReportDataList[i++] = new ReportEffectivenessDto { Theme = theme };
            }
            SetFormula();
        }

        public override AbstractReport CollectReportFromWs(string yymm)
        {
            var request = new GetReportRequest
            {
                Body = new GetReportRequestBody
                {
                    filialCode = FilialCode,
                    yymm = yymm,
                    reportType = ReportType.Effective
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response as ReportEffectiveness;

        }

        public override void MapForAutoFill(AbstractReport report)
        {
            if (report == null)
            {
                return;
            }
            var inReport = report as ReportEffectiveness;

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
            if (Report.ReportDataList != null && Report.ReportDataList.Length > 0 && Report.Status.ToString() != "New")
            {
                foreach (DataGridViewRow row in Dgv.Rows)
                {
                    var rowNum = row.Cells[0].Value.ToString();
                    var data = Report.ReportDataList.SingleOrDefault(x => x.CodeRowNum.ToString() == rowNum);
                    if (data != null)
                    {
                        row.Cells[1].Value = data.full_name ?? " ";
                        row.Cells[2].Value = data.expert_busyness;
                        row.Cells[3].Value = data.expert_speciality ?? " ";
                        row.Cells[4].Value = data.expertise_type ?? " ";
                        row.Cells[5].Value = data.mee_quantity_plan;
                        row.Cells[6].Value = data.mee_quantity_fact;
                        row.Cells[7].Value = data.mee_quantity_percent;
                        row.Cells[8].Value = data.mee_yeild_plan;
                        row.Cells[9].Value = data.mee_yeild_fact;
                        row.Cells[10].Value = data.mee_yeild_percent;
                        row.Cells[11].Value = data.ekmp_quantity_plan;
                        row.Cells[12].Value = data.ekmp_quantity_fact;
                        row.Cells[13].Value = data.ekmp_quantity_percent;
                        row.Cells[14].Value = data.ekmp_yeild_plan;
                        row.Cells[15].Value = data.ekmp_yeild_fact;
                        row.Cells[16].Value = data.ekmp_yeild_percent;
                    }
                }
                SetFormula();
            }
        }

        protected override void FillReport(string form)
        {
            if (form == null)
            {
                return;
            }
            var reportDto = new List<ReportEffectivenessDto>();
            foreach (DataGridViewRow row in Dgv.Rows)
            {
                var data = new ReportEffectivenessDto
                {
                    CodeRowNum = row.Cells[0].Value.ToString(),
                    full_name = row.Cells[1].Value.ToString() ?? " ",
                    expert_busyness = GlobalUtils.TryParseDecimal(row.Cells[2].Value ?? 0),
                    expert_speciality = row.Cells[3].Value.ToString() ?? " ",
                    expertise_type = row.Cells[4].Value.ToString() ?? " ",
                    mee_quantity_plan = GlobalUtils.TryParseInt(row.Cells[5].Value ?? 0),
                    mee_quantity_fact = GlobalUtils.TryParseInt(row.Cells[6].Value ?? 0),
                    mee_quantity_percent = GlobalUtils.TryParseDecimal(row.Cells[7].Value ?? 0),
                    mee_yeild_plan = GlobalUtils.TryParseInt(row.Cells[8].Value ?? 0),
                    mee_yeild_fact = GlobalUtils.TryParseInt(row.Cells[9].Value ?? 0),
                    mee_yeild_percent = GlobalUtils.TryParseDecimal(row.Cells[10].Value ?? 0),
                    ekmp_quantity_plan = GlobalUtils.TryParseInt(row.Cells[11].Value ?? 0),
                    ekmp_quantity_fact = GlobalUtils.TryParseInt(row.Cells[12].Value ?? 0),
                    ekmp_quantity_percent = GlobalUtils.TryParseDecimal(row.Cells[13].Value ?? 0),
                    ekmp_yeild_plan = GlobalUtils.TryParseInt(row.Cells[14].Value ?? 0),
                    ekmp_yeild_fact = GlobalUtils.TryParseInt(row.Cells[15].Value ?? 0),
                    ekmp_yeild_percent = GlobalUtils.TryParseDecimal(row.Cells[16].Value ?? 0),
                };
                reportDto.Add(data);
            }
            Report.ReportDataList = reportDto.ToArray();
        }

        public override bool IsVisibleBtnDownloadExcel() => false;

        public override bool IsVisibleBtnHandle() => false;

        public override string ValidReport() { return ""; }

        public override void ToExcel(string filename, string filialName)
        {

            //var excel = new ExcelCreator(filename, ExcelForm.effectiveness, Report.Yymm, filialName, Client, FilialCode);
            //excel.CreateReport(Report, null);
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
                    reportType = ReportType.Effective
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportEffectiveness;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
        }

        public override void SaveReportDataSourceExcel()
        { }
        public override void SaveReportDataSourceHandle()
        { }

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
                    reportType = ReportType.Effective
                }
            };
            var response = Client.CollectSummaryReport(request);
            Report = response.Body.CollectSummaryReportResult as ReportEffectiveness;
            Report.IdType = IdReportType;
            Report.Yymm = yymmEnd;
        }


        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            var formsList = ThemesList.Select(x => x.Key).OrderBy(x => x).ToList();
            var index = formsList.IndexOf(form);
            var currentHeaders = _headers[index];
            CreateDgvColumnsForTheme(Dgv, 50, _headersMap[form], currentHeaders);

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

        public void SetFormula()
        {
            //foreach (int row in Dgv.Rows)
            //{
            //        if (Dgv.Rows[row].Cells[2].Value == null)
            //        { Dgv.Rows[row].Cells[2].Value = 0.0; }
            //        if (Dgv.Rows[row].Cells[7].Value == null)
            //        { Dgv.Rows[row].Cells[7].Value = 0.0; }
            //        if (Dgv.Rows[row].Cells[10].Value == null)
            //        { Dgv.Rows[row].Cells[10].Value = 0.0; }
            //        if (Dgv.Rows[row].Cells[13].Value == null)
            //        { Dgv.Rows[row].Cells[13].Value = 0.0; }
            //        if (Dgv.Rows[row].Cells[16].Value == null)
            //        { Dgv.Rows[row].Cells[16].Value = 0.0; }
            //}
        }
    }
}