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
    public class ReportOpedFinanceProcessor : AbstractReportProcessor<ReportOpedFinance>
    {

        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        private readonly string[] _columns = { "Наименование показателя", "№ строки", "Фактическое значение показателя", "Примечание" };
        public ReportOpedFinanceProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
          base(inClient, dgv, cmb, txtb, page,
              XmlFormTemplate.OpedFinance.GetDescription(),
              Log,
              ReportGlobalConst.ReportOpedFinance,
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


        public void CalculateCells()
        {
            try
            {
                var row1 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "1.");
                if(row1 != null)
                {
                    row1.Cells[2].Value = Dgv.Rows.Cast<DataGridViewRow>().Where(x => x.Cells[1].Value.ToString() == "1.1."  || x.Cells[1].Value.ToString() == "1.2." || x.Cells[1].Value.ToString() == "1.4.").Sum(x=> GlobalUtils.TryParseDecimal(x.Cells[2].Value));
                }
            }
            catch (Exception ex) { }
        }

        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status)
        {

        }
        public override void InitReport()
        {
            Report = new ReportOpedFinance { ReportDataList = new ReportOpedFinanceData[] { }, IdType = IdReportType };
        }
        public override bool IsVisibleBtnDownloadExcel()
        {
            return false;
        }
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
            var excel = new ExcelOpedFinanceCreator(filename, ExcelForm.opedFinance, Report.Yymm, filialName);
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
                var N = new DataGridViewTextBoxCell { Value = row.Num };
                var cellname = new DataGridViewTextBoxCell { Value = row.Name };
                dgvRow.Cells.Add(cellname);
                dgvRow.Cells.Add(N);
                int rowIndex = Dgv.Rows.Add(dgvRow);
            }

            var row1 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "1.");
            row1.Cells[2].Style.BackColor = Color.LightGray;
            row1.Cells[2].ReadOnly = true;



        }
        protected override void FillReport(string form)
        {

            if (form == null)
            {
                return;
            }

            var reportDto = new List<ReportOpedFinanceData>();

            string[] notSavedCells = { "1." };

            foreach (DataGridViewRow row in Dgv.Rows)
            {
                string rowNum = row.Cells[1].Value.ToString();

                if (!notSavedCells.Contains(rowNum))
                {

                    var data = new ReportOpedFinanceData
                    {
                        RowNum = row.Cells[1].Value.ToString(),
                        ValueFact = GlobalUtils.TryParseDecimal(row.Cells[2].Value),
                        Notes = row.Cells[3].Value == null ? "" : row.Cells[3].Value.ToString()

                    };
                    reportDto.Add(data);
                }
                else
                {
                    var data = new ReportOpedFinanceData
                    {
                        RowNum = row.Cells[1].Value.ToString(),
                        ValueFact = 0,
                        Notes = row.Cells[3].Value == null ? "" : row.Cells[3].Value.ToString()

                    };
                    reportDto.Add(data);
                }

            }

            Report.ReportDataList = reportDto.ToArray();
        }
    }
}
