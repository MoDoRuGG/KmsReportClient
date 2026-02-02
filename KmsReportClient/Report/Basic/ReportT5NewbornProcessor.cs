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
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.XML;
using KmsReportClient.Support;
using NLog;

namespace KmsReportClient.Report.Basic
{
    internal class ReportT5NewbornProcessor : AbstractReportProcessor<ReportT5Newborn>
    {
        StackedHeaderDecorator DgvRender;

        private readonly List<string> based = new List<string>
        {
            "Доля рынка (наша)\nна 01.01.2026",
            "Застраховано новорожденных\n(нарастающим итогом отчетного года),\nчел.",
            "Всего реестров счетов\nот МО по родам\n(нарастающим итогом отчетного года),\nчел.",
            "Доля застрахованных\nот реестров счетов, %",
            "Отклонение от реестров\nтекущего месяца, чел.",
        };

        private static readonly Logger Log = LogManager.GetCurrentClassLogger();
        public ReportT5NewbornProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary,
                                        DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
        base(inClient, dgv, cmb, txtb, page,
            XmlFormTemplate.T5Newborn.GetDescription(),
            Log,
            ReportGlobalConst.ReportT5Newborn,
            reportsDictionary)
        {
            DgvRender = new StackedHeaderDecorator(Dgv);
            InitReport();
        }

        public override AbstractReport CollectReportFromWs(string yymm)
        {
            var request = new GetReportRequest
            {
                Body = new GetReportRequestBody
                {
                    filialCode = FilialCode,
                    yymm = yymm,
                    reportType = ReportType.T5Newborn
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as ReportT5Newborn;
        }

        public override void FillDataGridView(string form)
        {

            var reportT5Newborn = Report.ReportDataList.FirstOrDefault(x => x.Theme == form);
            if (reportT5Newborn == null)
            {
                return;
            }

            if (reportT5Newborn.Data != null)
            {
                Dgv.Rows[0].Cells[0].Value = reportT5Newborn.Data.MarketShare;
                Dgv.Rows[0].Cells[1].Value = reportT5Newborn.Data.CountNewborn;
                Dgv.Rows[0].Cells[2].Value = reportT5Newborn.Data.CountMaterinityBills;
                Dgv.Rows[0].Cells[3].Value = reportT5Newborn.Data.CountMaterinityBills != 0
                    ? Math.Round((reportT5Newborn.Data.CountNewborn / reportT5Newborn.Data.CountMaterinityBills * 100), 2)
                    : 0;
                Dgv.Rows[0].Cells[4].Value = reportT5Newborn.Data.CountNewborn - reportT5Newborn.Data.CountMaterinityBills;
            }

            foreach (DataGridViewColumn col in Dgv.Columns)
            {
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
        }

        public override void SaveReportDataSourceExcel()
        { }
        public override void SaveReportDataSourceHandle()
        { }
        public void SetFormula()
        { }

        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource)
        {

        }

        public override void InitReport()
        {
            Report = new ReportT5Newborn { ReportDataList = new ReportT5NewbornDto[ThemesList.Count], IdType = IdReportType };
            int i = 0;
            foreach (var theme in ThemesList.Select(x => x.Key))
            {
                Report.ReportDataList[i++] = new ReportT5NewbornDto { Theme = theme };
                Console.WriteLine(FilialName);
            }
        }

        public override bool IsVisibleBtnDownloadExcel() => false;

        public override bool IsVisibleBtnHandle() => false;

        public override bool IsVisibleBtnSummary() => false;

        public override void MapForAutoFill(AbstractReport report)
        {
            if (report == null)
            {
                return;
            }
            var inReport = report as ReportT5Newborn;

            var index = Report.ReportDataList.ToList().FindIndex(x => x.Theme == Cmb.Text);
            var inTheme = inReport.ReportDataList.Single(x => x.Theme == Cmb.Text);
            Report.ReportDataList[index] = inTheme;
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
                    reportType = ReportType.T5Newborn
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportT5Newborn;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
        }

        public override void ToExcel(string filename, string filialName)
        {
            //var excel = new ExcelT5NewbornCreator(filename, ExcelForm.T5Newborn, Report.Yymm, filialName, Client, FilialCode);
            //excel.CreateReport(Report, null);
        }

        public override string ValidReport() { return ""; }

        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            Dgv.AllowUserToAddRows = false;
            Dgv.ColumnHeadersVisible = true;

            Dgv.Columns.Clear();
            Dgv.Rows.Clear();

            List<string> columns = null;

            columns = based;


            foreach (var clmn in columns)
            {
                var column = new DataGridViewTextBoxColumn
                {
                    HeaderText = clmn,
                    DataPropertyName = "Indicator",
                    Name = "Indicator",
                    SortMode = DataGridViewColumnSortMode.NotSortable,
                    DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Azure }
                };

                Dgv.Columns.Add(column);
            }

            Dgv.Rows.Add();


            Dgv.Columns[3].ReadOnly =
            Dgv.Columns[4].ReadOnly = true;
            Dgv.AutoSize = true;
            Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

        }
        protected override void FillReport(string form)
        {
            var reportT5Newborn = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            var row = Dgv.Rows[0];
            reportT5Newborn.Data = new ReportT5NewbornDataDto
            {
                MarketShare = GlobalUtils.TryParseDecimal(row.Cells[0].Value),
                CountNewborn = GlobalUtils.TryParseDecimal(row.Cells[1].Value),
                CountMaterinityBills = GlobalUtils.TryParseDecimal(row.Cells[2].Value)
            };
        }
    }
}
