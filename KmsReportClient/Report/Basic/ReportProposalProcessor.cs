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
    public class ReportProposalProcessor : AbstractReportProcessor<ReportProposal>
    {


        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        private string[] _columns = new string[]
        {
            "Количество МО, подвергшихся проверкам в отчётном периоде, шт \n(1)",
            "Количество МО, подвергшихся проверкам в отчётном периоде, с выявленными нарушениями, шт\n(2)",
            "Количество предложений (писем), направленных СМО, шт\n(3)",
            "Количество МО, дефекты по которым отражены в предложениях (письмах), направленных СМО, шт\n(4)",
            "% выполнения (соотношение количества направленных СМО писем в ТФОМС к количеству МО с выявленными нарушениями)\n(5)",
            "Примечание\n(6)"
        };

        public ReportProposalProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
        base(inClient, dgv, cmb, txtb, page,
            XmlFormTemplate.Proposal.GetDescription(),
            Log,
            ReportGlobalConst.ReportProposal,
            reportsDictionary)
        {
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
                    reportType = ReportType.Proposal
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as ReportProposal;
        }
        public override void FillDataGridView(string form)
        {
            if (Report == null)
            {
                return;
            }

            Dgv.Rows[0].Cells[0].Value = Report.CountMoCheck;
            Dgv.Rows[0].Cells[1].Value = Report.CountMoCheckWithDefect;
            Dgv.Rows[0].Cells[2].Value = Report.CountProporsals;
            Dgv.Rows[0].Cells[3].Value = Report.CountProporsalsWithDefect;
            Dgv.Rows[0].Cells[5].Value = Report.Notes;

            CalculateCells();
        }
        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource)
        {

        }
        public override void InitReport()
        {
            Report = new ReportProposal { IdType = IdReportType };
        }
        public override bool IsVisibleBtnDownloadExcel() => false;

        public override bool IsVisibleBtnHandle() => false;

        public override void MapForAutoFill(AbstractReport report)
        {
            var inReport = report as ReportProposal;

            Report.IdReportData = inReport.IdReportData;

            Report.CountMoCheck = inReport.CountMoCheck;
            Report.CountMoCheckWithDefect = inReport.CountMoCheckWithDefect;
            Report.CountProporsals = inReport.CountProporsals;
            Report.CountProporsalsWithDefect = inReport.CountProporsalsWithDefect;
            Report.Notes = inReport.Notes;

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
                    reportType = ReportType.Proposal
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportProposal;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
        }
        public override void ToExcel(string filename, string filialName)
        {
            var excel = new ExcelProposalCreator(filename, ExcelForm.proposal, "", filialName, Report.Yymm);
            excel.CreateReport(Report, null);

        }

        public override void SaveReportDataSourceExcel()
        { }
        public override void SaveReportDataSourceHandle()
        { }

        public override string ValidReport()
        {
            return "";
        }
        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            Dgv.AllowUserToAddRows = false;
            Dgv.ColumnHeadersVisible = true;

            Dgv.Columns.Clear();
            Dgv.Rows.Clear();

            foreach (var clmn in _columns)
            {
                var column = new DataGridViewTextBoxColumn
                {
                    HeaderText = clmn,
                    DataPropertyName = "Indicator",
                    Name = "Indicator",
                    SortMode = DataGridViewColumnSortMode.NotSortable,
                    DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Azure },
                    Width = 150
                };

                Dgv.Columns.Add(column);
            }

            Dgv.Rows.Add();

            //Расчётный столбец
            Dgv.Columns[4].DefaultCellStyle.BackColor = Color.LightGray;
            Dgv.Columns[4].ReadOnly = true;


        }

        public void CalculateCells()
        {
            try
            {
                var row = Dgv.Rows[0];
                double res = (double)(GlobalUtils.TryParseDecimal(row.Cells[3].Value) / GlobalUtils.TryParseDecimal(row.Cells[1].Value)) * 100;
                row.Cells[4].Value = Math.Round(res, 2);
            }
            catch (Exception ex) { }
        }
        protected override void FillReport(string form)
        {
            var row = Dgv.Rows[0];
            Report.CountMoCheck = GlobalUtils.TryParseInt(row.Cells[0].Value);
            Report.CountMoCheckWithDefect = GlobalUtils.TryParseInt(row.Cells[1].Value);
            Report.CountProporsals = GlobalUtils.TryParseInt(row.Cells[2].Value);
            Report.CountProporsalsWithDefect = GlobalUtils.TryParseInt(row.Cells[3].Value);
            Report.Notes = row.Cells[5].Value == null ? "" : row.Cells[5].Value.ToString();
        }
    }
}
