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
    public class ReportInfrormationResponseProcessor : AbstractReportProcessor<ReportInfrormationResponse>
    {
        StackedHeaderDecorator DgvRender;

        private readonly List<string> profColumns = new List<string>
        {
            "Id",
            "Профилактические медосмотры  плановое значение   на 2021г.(человек)",
            "Текущий месяц;Проинформировано (человек)",
            "Текущий месяц;% к годовому плану",
            "Текущий месяц;Кол-во прошедших всего (по реестрам счетов)",
            "Текущий месяц;Кол-во прошедших ЗЛ из числа проинформированных",
            "Текущий месяц;%Отклик % (прошедшие из проинформированных)",
            "Профилактические медосмотры  плановое значение на 2021г.(человек)",
            "Исполнено с начала года;Проинформировано (человек)",
            "Исполнено с начала года;% к годовому плану",
            "Исполнено с начала года;Кол-во прошедших всего (по реестрам счетов)",
            "Исполнено с начала года;Кол-во прошедших ЗЛ из числа проинформированных",
            "Исполнено с начала года;Отклик % (прошедшие из проинформированных)",

        };

        private readonly List<string> dispColumns = new List<string>
        {
             "Id",
            "Диспансеризация плановое значение на 2021г.(человек)",
            "Текущий месяц;Проинформировано (человек)",
            "Текущий месяц;% к годовому плану",
            "Текущий месяц;Кол-во прошедших всего (по реестрам счетов)",
            "Текущий месяц;Кол-во прошедших ЗЛ из числа проинформированных",
            "Текущий месяц;%Отклик % (прошедшие из проинформированных)",
            "Диспансеризация  плановое значение   на 2021г.(человек)",
            "Исполнено с начала года;Проинформировано (человек)",
            "Исполнено с начала года;% к годовому плану",
            "Исполнено с начала года;Кол-во прошедших всего (по реестрам счетов)",
            "Исполнено с начала года;Кол-во прошедших ЗЛ из числа проинформированных",
            "Исполнено с начала года;Отклик % (прошедшие из проинформированных)",

        };

        private readonly List<string> dispNabColumns = new List<string>
        {
             "Id",
            "Диспансерное наблюдение плановое значение на 2021г.(человек)",
            "Текущий месяц;Проинформировано (человек)",
            "Текущий месяц;% к годовому плану",
            "Текущий месяц;Кол-во прошедших всего (по реестрам счетов)",
            "Текущий месяц;Кол-во прошедших ЗЛ из числа проинформированных",
            "Текущий месяц;%Отклик % (прошедшие из проинформированных)",
            "Диспансерное наблюдение плановое значение на 2021г.(человек)",
            "Исполнено с начала года;Проинформировано (человек)",
            "Исполнено с начала года;% к годовому плану",
            "Исполнено с начала года;Кол-во прошедших всего (по реестрам счетов)",
            "Исполнено с начала года;Кол-во прошедших ЗЛ из числа проинформированных",
            "Исполнено с начала года;Отклик % (прошедшие из проинформированных)",

        };

        private static readonly Logger Log = LogManager.GetCurrentClassLogger();
        public ReportInfrormationResponseProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
        base(inClient, dgv, cmb, txtb, page,
            XmlFormTemplate.FCR.GetDescription(),
            Log,
            ReportGlobalConst.ReportOtklik,
            reportsDictionary)
        {
            DgvRender = new StackedHeaderDecorator(Dgv);
            InitReport();


        }
        public override void SaveReportDataSourceExcel()
        { }
        public override void SaveReportDataSourceHandle()
        { }
        public override AbstractReport CollectReportFromWs(string yymm)
        {
            var request = new GetReportRequest
            {
                Body = new GetReportRequestBody
                {
                    filialCode = FilialCode,
                    yymm = yymm,
                    reportType = ReportType.IR
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as ReportInfrormationResponse;
            
        }
        public override void FillDataGridView(string form)
        {
            var reportInfrormation = Report.ReportDataList.FirstOrDefault(x => x.Theme == form);
            if (reportInfrormation == null)
            {
                return;
            }

            if (reportInfrormation.Data != null)
            {
                Dgv.Rows[0].Cells[0].Value = reportInfrormation.Data.Id;
                Dgv.Rows[0].Cells[1].Value = reportInfrormation.Data.Plan;
                Dgv.Rows[0].Cells[2].Value = reportInfrormation.Data.Informed;
                Dgv.Rows[0].Cells[4].Value = reportInfrormation.Data.CountPast;
                Dgv.Rows[0].Cells[5].Value = reportInfrormation.Data.CountRegistry;

            }

            var yearThemeData = Client.GetIRYearData(new GetIRYearDataRequest(new GetIRYearDataRequestBody
            {
                fillial = FilialCode,
                theme = form,
                yymm = Report.Yymm
            })).Body.GetIRYearDataResult;

            if (yearThemeData != null)
            {
                Dgv.Rows[0].Cells[7].Value = yearThemeData.Plan;
                Dgv.Rows[0].Cells[8].Value = yearThemeData.Informed;
                Dgv.Rows[0].Cells[10].Value = yearThemeData.CountPast;
                Dgv.Rows[0].Cells[11].Value = yearThemeData.CountRegistry;
            }
            SetFormula();


        }


        public void SetFormula()
        {
            //Console.WriteLine(Dgv.Rows[0].Cells[1].Value);
            //Console.WriteLine(Dgv.Rows[0].Cells[2].Value);
            //Console.WriteLine(GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[2].Value) / GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[1].Value));

            try
            {
              
                Dgv.Rows[0].Cells[3].Value = Math.Round(GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[2].Value) / GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[1].Value) * 100,2);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            try
            {
                Dgv.Rows[0].Cells[6].Value = Math.Round(GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[5].Value) / GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[2].Value) * 100, 2)  ;
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            Dgv.Rows[0].Cells[7].Value = Dgv.Rows[0].Cells[1].Value;

            try
            {
                Dgv.Rows[0].Cells[9].Value = Math.Round(GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[8].Value) / GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[7].Value) * 100, 2) ;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            try
            {
                Dgv.Rows[0].Cells[12].Value = Math.Round(GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[11].Value) / GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[8].Value) * 100, 2) ;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


            }


        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource)
        {

        }
        public override void InitReport()
        {
            Report = new ReportInfrormationResponse { ReportDataList = new ReportInfrormationResponseDto[ThemesList.Count], IdType = IdReportType };
            int i = 0;
            foreach (var theme in ThemesList.Select(x => x.Key))
            {
                Report.ReportDataList[i++] = new ReportInfrormationResponseDto { Theme = theme };
            }
        }
        public override bool IsVisibleBtnDownloadExcel() => true;

        public override bool IsVisibleBtnHandle() => false;

        public override bool IsVisibleBtnSummary() => false;

        public override void MapForAutoFill(AbstractReport report)
        {
            if (report == null)
            {
                return;
            }
            var inReport = report as ReportInfrormationResponse;

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
                    reportType = ReportType.IR
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportInfrormationResponse;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
        }
        public override void ToExcel(string filename, string filialName)
        {
            var excel = new ExceIRCreator(filename, ExcelForm.IR, Report.Yymm, filialName,Client,FilialCode);
            excel.CreateReport(Report, null);
        }
        public override string ValidReport() { return ""; }
        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            Dgv.AllowUserToAddRows = false;
            Dgv.ColumnHeadersVisible = true;

            Dgv.Columns.Clear();
            Dgv.Rows.Clear();

            List<string> columns = null;
            if (form == "профосмотр.")
            {
                columns = profColumns;
            }
            else if (form == "дисп.")
            {
                columns = dispColumns;

            }
            else if (form == "дисп.наблюд.")
            {
                columns = dispNabColumns;
            }

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
            Dgv.Columns[0].Visible = false;
            Dgv.Columns[1].Width = Dgv.Columns[7].Width = 140;
            Dgv.Columns[2].Width = Dgv.Columns[8].Width = 120;
            Dgv.Columns[3].Width = Dgv.Columns[9].Width = 150;
            Dgv.Columns[4].Width = Dgv.Columns[10].Width = 150;
            Dgv.Columns[5].Width = Dgv.Columns[11].Width = 150;
            Dgv.Columns[6].Width = Dgv.Columns[12].Width = 150;

            Dgv.Columns[3].DefaultCellStyle.BackColor =
            Dgv.Columns[6].DefaultCellStyle.BackColor =
            Dgv.Columns[7].DefaultCellStyle.BackColor =
            Dgv.Columns[7].DefaultCellStyle.BackColor =
            Dgv.Columns[8].DefaultCellStyle.BackColor =
            Dgv.Columns[9].DefaultCellStyle.BackColor =
            Dgv.Columns[10].DefaultCellStyle.BackColor =
            Dgv.Columns[11].DefaultCellStyle.BackColor =
            Dgv.Columns[12].DefaultCellStyle.BackColor = Color.LightGray;


            Dgv.Columns[3].ReadOnly =
            Dgv.Columns[6].ReadOnly =
            Dgv.Columns[7].ReadOnly =
            Dgv.Columns[7].ReadOnly =
            Dgv.Columns[8].ReadOnly =
            Dgv.Columns[9].ReadOnly =
            Dgv.Columns[10].ReadOnly =
            Dgv.Columns[11].ReadOnly =
            Dgv.Columns[12].ReadOnly = true;


        }
        protected override void FillReport(string form)
        {
            var reportInfrormation = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            var row = Dgv.Rows[0];
            reportInfrormation.Data = new ReportInfrormationResponseDataDto
            {
                Id = GlobalUtils.TryParseInt(row.Cells[0].Value),
                Plan = GlobalUtils.TryParseInt(row.Cells[1].Value),
                Informed = GlobalUtils.TryParseInt(row.Cells[2].Value),
                CountPast = GlobalUtils.TryParseInt(row.Cells[4].Value),
                CountRegistry = GlobalUtils.TryParseInt(row.Cells[5].Value)
            };



        }
    }
}
