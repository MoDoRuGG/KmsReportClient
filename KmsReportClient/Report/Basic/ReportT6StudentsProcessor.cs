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
    internal class ReportT6StudentsProcessor : AbstractReportProcessor<ReportT6Students>
    {
        StackedHeaderDecorator DgvRender;

        private readonly List<string> based = new List<string>
        {
            "Количество учебных заведений по региону (ВУЗы), шт.",
            "Количество учебных заведений по региону (Колледжи и пр.уч.заведения), шт.",
            "Застраховано за текущий календарный год (с января нарастающим итогом) чел.",
            "\"Комментарии по сотрудничеству с учебными заведениями\r\n1. количество УЗ с которыми сотрудничаем\r\n2. наименование УЗ\r\n3. каким образом организовано сотрудничество \"\r\n",
        };

        private static readonly Logger Log = LogManager.GetCurrentClassLogger();
        public ReportT6StudentsProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary,
                                        DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
        base(inClient, dgv, cmb, txtb, page,
            XmlFormTemplate.T6Students.GetDescription(),
            Log,
            ReportGlobalConst.ReportT6Students,
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
                    reportType = ReportType.T6Students
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as ReportT6Students;
        }

        public override void FillDataGridView(string form)
        {

            var reportT6Students = Report.ReportDataList.FirstOrDefault(x => x.Theme == form);
            if (reportT6Students == null)
            {
                return;
            }

            if (reportT6Students.Data != null)
            {
                Dgv.Rows[0].Cells[0].Value = reportT6Students.Data.CountUniversity;
                Dgv.Rows[0].Cells[1].Value = reportT6Students.Data.CountCollege;
                Dgv.Rows[0].Cells[2].Value = reportT6Students.Data.CountInsured;
                Dgv.Rows[0].Cells[3].Value = reportT6Students.Data.Comments;
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
            Report = new ReportT6Students { ReportDataList = new ReportT6StudentsDto[ThemesList.Count], IdType = IdReportType };
            int i = 0;
            foreach (var theme in ThemesList.Select(x => x.Key))
            {
                Report.ReportDataList[i++] = new ReportT6StudentsDto { Theme = theme };
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
            var inReport = report as ReportT6Students;

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
                    reportType = ReportType.T6Students
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportT6Students;
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
            Dgv.AutoSize = true;
            Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

        }
        protected override void FillReport(string form)
        {
            var reportT6Students = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            var row = Dgv.Rows[0];
            reportT6Students.Data = new ReportT6StudentsDataDto
            {
                CountUniversity = GlobalUtils.TryParseInt(row.Cells[0].Value),
                CountCollege = GlobalUtils.TryParseInt(row.Cells[1].Value),
                CountInsured = GlobalUtils.TryParseInt(row.Cells[2].Value),
                Comments = row.Cells[2].Value.ToString()
            };
        }
    }
}
