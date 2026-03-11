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
using KmsReportClient.Model;
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.XML;
using KmsReportClient.Support;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.OLE.Interop;
using NLog;

namespace KmsReportClient.Report.Basic
{
    public class ReportPVPLoadProcessor : AbstractReportProcessor<ReportPVPLoad>
    {
        private readonly List<string> headers = new List<string>
        {
            "Наименование ПВП\r\n",
            "место размещения офиса/МП филиала + адрес\r\n",
            "Численность застрахованных филиалом на начало периода\r\n(1)",
            "Численность застрахованных филиалом на отчетную дату\r\n(2)",
            "Динамика численности (нарастающим итогом за текущий год)\r\n(3)",
            "Специалист (Ф.И.О.)\r\n(4)",
            "условия трудоустройства (размер ставки)\r\n(5)",
            "План ПВП по вновь застрахованным на год., чел.\r\n(6)",
            "оформлено всего граждан (новых, переоформление, перерегистрация)\r\n(7)",
            "В т.ч. вновь застрахованных, чел.\r\n(8)",
            "В т.ч. кол-во застрахованных, привлеченных агентами\r\n(9)",
            "выдано ПЕО и выписок из ЕРЗЛ\r\n(10)",
            "Всего обслужено населения\nгр. 7 + гр.10\r\n(11)",
            "отклонения от плана\n(гр.8-гр.6)\r\n(12)",
            "Нагрузка в день на 1 спец-та (гр. 11/кол-во раб. дней)*гр.5)\r\n(13)",
            "Обращений через госуслуги\r\n(14)",
            "Примечание\r\n"



        };
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();
        public ReportPVPLoadProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, System.Windows.Forms.TextBox txtb, TabPage page) :
        base(inClient, dgv, cmb, txtb, page,
            XmlFormTemplate.PVPL.GetDescription(),
            Log,
            ReportGlobalConst.ReportPVPLoad,
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
                    reportType = ReportType.PVPLoad
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as ReportPVPLoad;
        }

        public override void FillDataGridView(string form)
        {
            if (Report != null)
            {
                if (Report.Data != null)
                {
                    foreach (DataGridViewRow row in Dgv.Rows)
                    {
                        var rowData = Report.Data.FirstOrDefault(x => x.RowNumID == row.Index);
                       if (rowData != null)
                        {

                            row.Cells[0].Value = rowData.PVP_name;
                            row.Cells[1].Value = rowData.location_of_the_office;
                            row.Cells[2].Value = rowData.number_of_insured_by_beginning_of_year;
                            row.Cells[3].Value = rowData.number_of_insured_by_reporting_date;
                            row.Cells[4].Value = rowData.population_dynamics;
                            row.Cells[5].Value = rowData.specialist;
                            row.Cells[6].Value = rowData.conditions_of_employment;
                            row.Cells[7].Value = rowData.PVP_plan;
                            row.Cells[8].Value = rowData.registered_total_citizens;
                            row.Cells[9].Value = rowData.newly_insured;
                            row.Cells[10].Value = rowData.attracted_by_agents;
                            row.Cells[11].Value = rowData.issued_by_PEO_and_extracts_from_ERZL;
                            row.Cells[12].Value = rowData.registered_total_citizens + rowData.issued_by_PEO_and_extracts_from_ERZL;
                            row.Cells[13].Value = rowData.newly_insured - rowData.PVP_plan;
                            row.Cells[14].Value = rowData.workload_per_day_for_specialist;
                            row.Cells[15].Value = rowData.appeals_through_EPGU;
                            row.Cells[16].Value = rowData.notes;

                        }
                    }
                }
            }
        }
        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource)
        {

        }
        public override void InitReport()
        {
            Report = new ReportPVPLoad { IdType = IdReportType };
        }
        public override bool IsVisibleBtnDownloadExcel() => true;

        public override bool IsVisibleBtnHandle() => false;

        public override bool IsVisibleBtnSummary() => false;

        public override void MapForAutoFill(AbstractReport report)
        {
            var inReport = report as ReportPVPLoad;
            if (inReport?.Data != null)
            {
                Report.Data = inReport.Data.Select(x => new PVPload
                {
                    RowNumID = x.RowNumID,
                    PVP_name = x.PVP_name,
                    location_of_the_office = x.location_of_the_office,
                    number_of_insured_by_beginning_of_year = x.number_of_insured_by_beginning_of_year,
                    number_of_insured_by_reporting_date = x.number_of_insured_by_reporting_date,
                    population_dynamics = x.population_dynamics,
                    specialist = x.specialist,
                    conditions_of_employment = x.conditions_of_employment,
                    PVP_plan = x.PVP_plan,
                    registered_total_citizens = x.registered_total_citizens,
                    newly_insured = x.newly_insured,
                    attracted_by_agents = x.attracted_by_agents,
                    issued_by_PEO_and_extracts_from_ERZL = x.issued_by_PEO_and_extracts_from_ERZL,
                    workload_per_day_for_specialist = x.workload_per_day_for_specialist,
                    appeals_through_EPGU = x.appeals_through_EPGU,
                    notes = x.notes
                }).ToArray();
            }
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
                    reportType = ReportType.PVPLoad
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportPVPLoad;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
        }
        public override void ToExcel(string filename, string filialName)
        {
            var excel = new ExcelPVPLoadCreator(filename, ExcelForm.PVPLoad, Report.Yymm, filialName, Client, FilialCode);
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
            Dgv.AllowUserToAddRows = true;
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

            int RowCounter = Report.Data == null ? 0 : Report.Data.Length;
            if (RowCounter > 0)
            {
                for (int i = 0; i < RowCounter; i++)
                    Dgv.Rows.Add();
            }
            else
            {
                Dgv.Rows.Add();
            }

        }

        public void SetFormula(int workDays = 21)
        {
            foreach (DataGridViewRow row in Dgv.Rows)
            {
                if (row.IsNewRow) continue;

                // Входные данные
                var rate = GlobalUtils.TryParseDecimal(row.Cells[6].Value);      // гр.5
                var plan = GlobalUtils.TryParseInt(row.Cells[7].Value);          // гр.6
                var totalCitizens = GlobalUtils.TryParseInt(row.Cells[8].Value); // гр.7
                var newlyInsured = GlobalUtils.TryParseInt(row.Cells[9].Value);  // гр.8
                var peo = GlobalUtils.TryParseInt(row.Cells[11].Value);          // гр.10

                // Расчеты
                row.Cells[12].Value = totalCitizens + peo;                       // гр.11 (7+10)
                row.Cells[13].Value = newlyInsured - plan;                       // гр.12 (8-6)
                row.Cells[12].Style.BackColor = row.Cells[13].Style.BackColor = Color.LightGray;

                // Блокировка редактирования
                for (int i = 12; i <= 13; i++) row.Cells[i].ReadOnly = true;
            }
        }

        protected override void FillReport(string form)
        {
            List<PVPload> dataList = new List<PVPload>();
            foreach (DataGridViewRow row in Dgv.Rows)
            {
                if (row.Index + 1 < Dgv.Rows.Count)
                {
                    int rowNum = row.Index;
                    dataList.Add(new PVPload
                    {
                        RowNumID = row.Index,
                        PVP_name = row.Cells[0].Value?.ToString() ?? "",
                        location_of_the_office = row.Cells[1].Value?.ToString() ?? "",
                        number_of_insured_by_beginning_of_year = GlobalUtils.TryParseInt(row.Cells[2].Value),
                        number_of_insured_by_reporting_date = GlobalUtils.TryParseInt(row.Cells[3].Value),
                        population_dynamics = GlobalUtils.TryParseInt(row.Cells[4].Value),
                        specialist = row.Cells[5].Value?.ToString() ?? "",
                        conditions_of_employment = GlobalUtils.TryParseDecimal(row.Cells[6].Value),
                        PVP_plan = GlobalUtils.TryParseInt(row.Cells[7].Value),
                        registered_total_citizens = GlobalUtils.TryParseInt(row.Cells[8].Value),
                        newly_insured = GlobalUtils.TryParseInt(row.Cells[9].Value),
                        attracted_by_agents = GlobalUtils.TryParseInt(row.Cells[10].Value),
                        issued_by_PEO_and_extracts_from_ERZL = GlobalUtils.TryParseInt(row.Cells[11].Value), // Исправлено: было 10
                                                                                                             // Ячейки 12 и 13 расчетные, не сохраняем
                        workload_per_day_for_specialist = GlobalUtils.TryParseDecimal(row.Cells[14].Value), // Исправлено: было 12
                        appeals_through_EPGU = GlobalUtils.TryParseInt(row.Cells[15].Value), // Исправлено: было 13
                        notes = row.Cells[16].Value?.ToString() ?? "" // Исправлено: было 14
                    });

                    Report.Data = dataList.ToArray();
                }
            }
        }
    }
}
