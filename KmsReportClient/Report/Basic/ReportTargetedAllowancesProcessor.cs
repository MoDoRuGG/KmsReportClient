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
using NLog;

namespace KmsReportClient.Report.Basic
{
    public class ReportTargetedAllowancesProcessor : AbstractReportProcessor<ReportTargetedAllowances>
    {
        private readonly List<string> headers = new List<string>
        {
            "ФИО",
            "Специальность",
            "Период (месяц)",
            "Количество ЭКМП",
            "Сумма удержаний",
            "Сумма оплаты эксперту",
            "Предоставивший эксперта филиал",
            "Комментарии"
        };
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();
        public ReportTargetedAllowancesProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
        base(inClient, dgv, cmb, txtb, page,
            XmlFormTemplate.TarAllow.GetDescription(),
            Log,
            ReportGlobalConst.ReportTargetedAllowances,
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
                    reportType = ReportType.TarAllow
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as ReportTargetedAllowances;
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
                            row.Cells[0].Value = rowData.FIO;
                            row.Cells[1].Value = rowData.Speciality;
                            row.Cells[2].Value = rowData.Period;
                            row.Cells[3].Value = rowData.CountEKMP;
                            row.Cells[4].Value = rowData.AmountSank;
                            row.Cells[5].Value = rowData.AmountPayment;
                            row.Cells[6].Value = rowData.ProvidedBy;
                            row.Cells[7].Value = rowData.Comments;
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
            Report = new ReportTargetedAllowances { IdType = IdReportType };
        }
        public override bool IsVisibleBtnDownloadExcel() => false;

        public override bool IsVisibleBtnHandle() => false;

        public override bool IsVisibleBtnSummary() => false;

        public override void MapForAutoFill(AbstractReport report)
        {
            var inReport = report as ReportTargetedAllowances;
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
                    reportType = ReportType.TarAllow
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportTargetedAllowances;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
        }
        public override void ToExcel(string filename, string filialName)
        {
            var excel = new ExcelTarAllowCreator(filename, ExcelForm.TarAllow, Report.Yymm, filialName, Client, FilialCode);
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

        public void SetFormula()
        {
        }
        protected override void FillReport(string form)
        {
            List<TargetedAllowancesData> dataList = new List<TargetedAllowancesData>();
            foreach (DataGridViewRow row in Dgv.Rows)
            {
                if (row.Index + 1 < Dgv.Rows.Count)
                {
                    int rowNum = row.Index;
                    dataList.Add(new TargetedAllowancesData
                    {
                        RowNumID = rowNum,
                        FIO = row.Cells[0].Value == null ? "" : row.Cells[0].Value.ToString(),
                        Speciality = row.Cells[1].Value == null ? "" : row.Cells[1].Value.ToString(),
                        Period = row.Cells[2].Value == null ? "" : row.Cells[2].Value.ToString(),
                        CountEKMP = GlobalUtils.TryParseInt(row.Cells[3].Value),
                        AmountSank = GlobalUtils.TryParseDecimal(row.Cells[4].Value),
                        AmountPayment = GlobalUtils.TryParseDecimal(row.Cells[5].Value),
                        ProvidedBy = row.Cells[6].Value == null ? "" : row.Cells[6].Value.ToString(),
                        Comments = row.Cells[7].Value == null ? "" : row.Cells[7].Value.ToString()
                    });


                    Report.Data = dataList.ToArray();
                }
            }
        }
    }
}
