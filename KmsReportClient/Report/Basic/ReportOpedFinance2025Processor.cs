using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
    public class ReportOpedFinance2025Processor : AbstractReportProcessor<ReportOpedFinance>
    {

        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        private readonly string[] _columns = { "Наименование показателя", "№ п/п", "Фактическое значение показателя (руб.)", "Примечание" };
        public ReportOpedFinance2025Processor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
          base(inClient, dgv, cmb, txtb, page,
              XmlFormTemplate.OpedFinance2025.GetDescription(),
              Log,
              ReportGlobalConst.ReportOpedFinance2025,
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
                    reportType = ReportType.OpedFinance
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as ReportOpedFinance;
        }
        public override void FillDataGridView(string form)
        {
            if (form == null)
            {
                return;
            }

            if (Report.ReportDataList != null && Report.ReportDataList.Length > 0)
            {

                foreach (DataGridViewRow row in Dgv.Rows)
                {
                    var rowNum = row.Cells[1].Value.ToString();

                    var data = Report.ReportDataList.SingleOrDefault(x => x.RowNum.ToString() == rowNum);


                    if (data != null)
                    {
                        row.Cells[2].Value = data.ValueFact;
                        row.Cells[3].Value = data.Notes;


                    }
                }

                CalculateCells();

            }
        }

        public override void SaveReportDataSourceExcel()
        { }
        public override void SaveReportDataSourceHandle()
        { }

        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource)
        {

        }
        public override void InitReport()
        {
            Report = new ReportOpedFinance { ReportDataList = new ReportOpedFinanceData[] { }, IdType = IdReportType };
        }
        public override bool IsVisibleBtnDownloadExcel() => false;

        public override bool IsVisibleBtnHandle() => false;
        public override bool IsVisibleBtnSummary() => false;

        public override void MapForAutoFill(AbstractReport report)
        {

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
                    reportType = ReportType.OpedFinance
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportOpedFinance;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
        }
        public override void ToExcel(string filename, string filialName)
        {
            var excel = new ExcelOpedFinance2025Creator(filename, ExcelForm.OpedFinance2025, Report.Yymm, filialName);
            excel.CreateReport(Report, null);
        }
        public override string ValidReport()
        {
            return "";
        }
        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            Dgv.AllowUserToAddRows = false;
            Dgv.ColumnHeadersVisible = true;

            foreach (var col in _columns)
            {
                var dgvColumn = new DataGridViewTextBoxColumn
                {
                    HeaderText = col,
                    Width = 150,
                    ReadOnly = false,
                    SortMode = DataGridViewColumnSortMode.NotSortable
                };

                Dgv.Columns.Add(dgvColumn);
            }

            foreach (var row in table)
            {
                var dgvRow = new DataGridViewRow();
                var N = new DataGridViewTextBoxCell { Value = row.RowNum_fromxml };
                var cellname = new DataGridViewTextBoxCell { Value = row.RowText_fromxml };
                dgvRow.Cells.Add(cellname);
                dgvRow.Cells.Add(N);
                int rowIndex = Dgv.Rows.Add(dgvRow);
            }
        }
        protected override void FillReport(string form)
        {

            if (form == null)
            {
                return;
            }

            var reportDto = new List<ReportOpedFinanceData>();

            foreach (DataGridViewRow row in Dgv.Rows)
            {
                    var data = new ReportOpedFinanceData
                    {
                        RowNum = row.Cells[1].Value.ToString(),
                        ValueFact = GlobalUtils.TryParseDecimal(row.Cells[2].Value),
                        Notes = row.Cells[3].Value == null ? "" : row.Cells[3].Value.ToString()

                    };
                    reportDto.Add(data);
            }
            Report.ReportDataList = reportDto.ToArray();
        }
    }
}
