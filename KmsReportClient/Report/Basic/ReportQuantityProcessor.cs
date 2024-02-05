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
using KmsReportClient.Forms;
using KmsReportClient.Global;
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.XML;
using KmsReportClient.Support;
using NLog;

namespace KmsReportClient.Report.Basic
{
    public class ReportQuantityProcessor : AbstractReportProcessor<ReportQuantity>
    {


        
        StackedHeaderDecorator DgvRender;
        private readonly List<string> headers = new List<string>
        {
            "Численность, принятая к финансированию ТФОМС в предыдущий период",
            "Кол-во вновь застрахованных лиц в отчетном периоде, принятых к финансированию в ТФОМС;Всего",
            "Кол-во вновь застрахованных лиц в отчетном периоде, принятых к финансированию в ТФОМС;из них;принято ТФОМC",
            "Кол-во вновь застрахованных лиц в отчетном периоде, принятых к финансированию в ТФОМС;из них;не принято ТФОМС",
            "Кол-во лиц, из застрахованных филиалом, обратившихся в отчетном периоде за полисом ОМС",
            "Кол-во лиц подлежащих страхованию, сведения о которых получены от ТФОМС;Всего",
            "Кол-во лиц подлежащих страхованию, сведения о которых получены от ТФОМС;Кол-во лиц данной категории обратившихся за полисом ОМС в отчетном периоде",
            "Кол-во ранее застрахованных филиалом лиц, записи о которых исключены из финансирования ТФОМС в отчетном периоде и подлежат закрытию в собственном сегменте регистра филиала;Всего",
            "Кол-во ранее застрахованных филиалом лиц, записи о которых исключены из финансирования ТФОМС в отчетном периоде и подлежат закрытию в собственном сегменте регистра филиала;Из них по причинам;Смерть (код 1)",
            "Кол-во ранее застрахованных филиалом лиц, записи о которых исключены из финансирования ТФОМС в отчетном периоде и подлежат закрытию в собственном сегменте регистра филиала;Из них по причинам;Замены СМО на территории (код 2)",
            "Кол-во ранее застрахованных филиалом лиц, записи о которых исключены из финансирования ТФОМС в отчетном периоде и подлежат закрытию в собственном сегменте регистра филиала;Из них по причинам;Замены СМО в связи с изменениями места жительства (код 3)",
            "Кол-во ранее застрахованных филиалом лиц, записи о которых исключены из финансирования ТФОМС в отчетном периоде и подлежат закрытию в собственном сегменте регистра филиала;Из них по причинам;Выбор другой СМО (код 4)",
            "Кол-во ранее застрахованных филиалом лиц, записи о которых исключены из финансирования ТФОМС в отчетном периоде и подлежат закрытию в собственном сегменте регистра филиала;Из них по причинам;Дубликаты ЕРЗ (код 5)",
            "Кол-во ранее застрахованных филиалом лиц, записи о которых исключены из финансирования ТФОМС в отчетном периоде и подлежат закрытию в собственном сегменте регистра филиала;Из них по причинам;Другие (код 6)",
            "Кол-во лиц записи о которых восстановлены в сегменте регистра филиала в отчетном периоде",
            "Численность принятая к финансированию ТФОМС в отчетном периоде"
        };
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();
        public ReportQuantityProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
        base(inClient, dgv, cmb, txtb, page,
            XmlFormTemplate.Quantity.GetDescription(),
            Log,
            ReportGlobalConst.ReportQuantity,
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
                    reportType = ReportType.Quantity
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as ReportQuantity;
        }

        public override void FillDataGridView(string form)
        {
            if (Report == null)
            {
                return;
            }

            Dgv.Rows[0].Cells[0].Value = Report.Col_1;
            Dgv.Rows[0].Cells[1].Value = Report.Col_2;
            Dgv.Rows[0].Cells[2].Value = Report.Col_3;
            Dgv.Rows[0].Cells[3].Value = Report.Col_4;
            Dgv.Rows[0].Cells[4].Value = Report.Col_5;
            Dgv.Rows[0].Cells[5].Value = Report.Col_6;
            Dgv.Rows[0].Cells[6].Value = Report.Col_7;
            Dgv.Rows[0].Cells[7].Value = Report.Col_8;
            Dgv.Rows[0].Cells[8].Value = Report.Col_9;
            Dgv.Rows[0].Cells[9].Value = Report.Col_10;
            Dgv.Rows[0].Cells[10].Value = Report.Col_11;
            Dgv.Rows[0].Cells[11].Value = Report.Col_12;
            Dgv.Rows[0].Cells[12].Value = Report.Col_13;
            Dgv.Rows[0].Cells[13].Value = Report.Col_14;
            Dgv.Rows[0].Cells[14].Value = Report.Col_15;
            Dgv.Rows[0].Cells[15].Value = Report.Col_16;

            //CalculateCells();
        }
        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource)
        {

        }
        public override void InitReport()
        {
            Report = new ReportQuantity { IdType = IdReportType };
        }
        public override bool IsVisibleBtnDownloadExcel() => false;

        public override bool IsVisibleBtnHandle() => false;

        public override bool IsVisibleBtnSummary() => false;

        public override void MapForAutoFill(AbstractReport report)
        {
            var inReport = report as ReportQuantity;

            Report.Id_Report_Data = inReport.Id_Report_Data;
            Report.Col_1 = inReport.Col_16;
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
                    reportType = ReportType.Quantity
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportQuantity;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
        }
        public override void ToExcel(string filename, string filialName)
        {
           var excel = new ExcelQuantityCreator(filename, ExcelForm.Quantity, Report.Yymm, filialName, Client, FilialCode);
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

            foreach (var clmn in headers)
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
            Dgv.Columns[1].DefaultCellStyle.BackColor = Color.LightGray;
            Dgv.Columns[1].ReadOnly = true;
            Dgv.Columns[7].DefaultCellStyle.BackColor = Color.LightGray;
            Dgv.Columns[7].ReadOnly = true;
            Dgv.Columns[15].DefaultCellStyle.BackColor = Color.LightGray;
            Dgv.Columns[15].ReadOnly = true;


        }

        public void SetFormula()
        {
            try
            {
                var row = Dgv.Rows[0];
                row.Cells[1].Value = Convert.ToInt32(row.Cells[2].Value) + Convert.ToInt32(row.Cells[3].Value);
                row.Cells[7].Value = Convert.ToInt32(row.Cells[8].Value) + Convert.ToInt32(row.Cells[9].Value) + Convert.ToInt32(row.Cells[10].Value) +
                                     Convert.ToInt32(row.Cells[11].Value) + Convert.ToInt32(row.Cells[12].Value) + Convert.ToInt32(row.Cells[13].Value);
                row.Cells[15].Value = Convert.ToInt32(row.Cells[0].Value) + Convert.ToInt32(row.Cells[2].Value) + Convert.ToInt32(row.Cells[6].Value) -
                                      Convert.ToInt32(row.Cells[7].Value) + Convert.ToInt32(row.Cells[14].Value);
            }
            catch (Exception ex) { }
        }
        protected override void FillReport(string form)
        {
            var row = Dgv.Rows[0];
            Report.Col_1 = GlobalUtils.TryParseInt(row.Cells[0].Value);
            Report.Col_2 = GlobalUtils.TryParseInt(row.Cells[1].Value);
            Report.Col_3 = GlobalUtils.TryParseInt(row.Cells[2].Value);
            Report.Col_4 = GlobalUtils.TryParseInt(row.Cells[3].Value);
            Report.Col_5 = GlobalUtils.TryParseInt(row.Cells[4].Value);
            Report.Col_6 = GlobalUtils.TryParseInt(row.Cells[5].Value);
            Report.Col_7 = GlobalUtils.TryParseInt(row.Cells[6].Value);
            Report.Col_8 = GlobalUtils.TryParseInt(row.Cells[7].Value);
            Report.Col_9 = GlobalUtils.TryParseInt(row.Cells[8].Value);
            Report.Col_10 = GlobalUtils.TryParseInt(row.Cells[9].Value);
            Report.Col_11 = GlobalUtils.TryParseInt(row.Cells[10].Value);
            Report.Col_12 = GlobalUtils.TryParseInt(row.Cells[11].Value);
            Report.Col_13 = GlobalUtils.TryParseInt(row.Cells[12].Value);
            Report.Col_14 = GlobalUtils.TryParseInt(row.Cells[13].Value);
            Report.Col_15 = GlobalUtils.TryParseInt(row.Cells[14].Value);
            Report.Col_16 = GlobalUtils.TryParseInt(row.Cells[15].Value);
        }
    }
}
