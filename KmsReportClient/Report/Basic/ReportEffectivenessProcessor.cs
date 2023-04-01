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
        ): 
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
            //SetFormula();
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
            var reportEffectivenessDto = Report.ReportDataList?.Single(x => x.Theme == form);
            if (reportEffectivenessDto.Data == null || reportEffectivenessDto.Data.Length == 0)
            {
                return;
            }

            var rows = ThemeTextData.tables.Where(x => x.Name == form).SelectMany(x => x.Rows).ToList();
            foreach (DataGridViewRow row in Dgv.Rows)
            {
                var rowNum = row.Cells[0].Value.ToString().Trim();
                var data = reportEffectivenessDto.Data.SingleOrDefault(x => x.CodeRowNum == rowNum);
                if (data == null)
                {
                    continue;
                }

                row.Cells[1].Value = data.full_name;
                row.Cells[2].Value = data.expert_busyness.ToString();
                row.Cells[3].Value = data.expert_speciality;
                row.Cells[4].Value = data.expertise_type;
                row.Cells[5].Value = data.mee_quantity_plan.ToString();
                row.Cells[6].Value = data.mee_quantity_fact.ToString();
                row.Cells[7].Value = data.mee_quantity_percent.ToString();
                row.Cells[8].Value = data.mee_yeild_plan.ToString();
                row.Cells[9].Value = data.mee_yeild_fact.ToString();
                row.Cells[10].Value = data.mee_yeild_percent.ToString();
                row.Cells[11].Value = data.ekmp_quantity_plan.ToString();
                row.Cells[12].Value = data.ekmp_quantity_fact.ToString();
                row.Cells[13].Value = data.ekmp_quantity_percent.ToString();
                row.Cells[14].Value = data.ekmp_yeild_plan.ToString();
                row.Cells[15].Value = data.ekmp_yeild_fact.ToString();
                row.Cells[16].Value = data.ekmp_yeild_percent.ToString(); 
            }
            //SetFormula();
        }

        protected override void FillReport(string form)
        {
            if (form == null)
            {
                return;
            }
            var reportEffectivenessDto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportEffectivenessDto != null)
            {
                reportEffectivenessDto.Data = (from DataGridViewRow row in Dgv.Rows
                                               let rowNum = row.Cells[0].Value.ToString().Trim()
                                               select new ReportEffectivenessDataDto
                                               {
                                                   CodeRowNum = rowNum,  
                                                   full_name = row.Cells[1].ToString(),
                                                   expert_busyness = GlobalUtils.TryParseDecimal(row.Cells[2].Value),
                                                   expert_speciality = row.Cells[3].Value.ToString(),
                                                   expertise_type = row.Cells[4].Value.ToString(),
                                                   mee_quantity_plan = GlobalUtils.TryParseInt(row.Cells[5].Value),
                                                   mee_quantity_fact = GlobalUtils.TryParseInt(row.Cells[6].Value),
                                                   mee_quantity_percent = GlobalUtils.TryParseDecimal(row.Cells[7].Value),
                                                   mee_yeild_plan = GlobalUtils.TryParseInt(row.Cells[8].Value),
                                                   mee_yeild_fact = GlobalUtils.TryParseInt(row.Cells[9].Value),
                                                   mee_yeild_percent = GlobalUtils.TryParseDecimal(row.Cells[10].Value),
                                                   ekmp_quantity_plan = GlobalUtils.TryParseInt(row.Cells[11].Value),
                                                   ekmp_quantity_fact = GlobalUtils.TryParseInt(row.Cells[12].Value),
                                                   ekmp_quantity_percent = GlobalUtils.TryParseDecimal(row.Cells[13].Value),
                                                   ekmp_yeild_plan = GlobalUtils.TryParseInt(row.Cells[14].Value),
                                                   ekmp_yeild_fact = GlobalUtils.TryParseInt(row.Cells[15].Value),
                                                   ekmp_yeild_percent = GlobalUtils.TryParseDecimal(row.Cells[16].Value),
                                               }).ToArray();
            }

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
            //SetFormula();
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
                        Value = row.Exclusion || isNeedExcludeSum ? "x" : "0"
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

            //try
            //{
            //    Dgv.Rows[0].Cells[3].Value = Math.Round(GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[6].Value) + GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[9].Value) +
            //                                            GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[12].Value), 2);
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            //}

        }
    }
}
