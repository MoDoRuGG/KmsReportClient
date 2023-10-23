using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using Org.BouncyCastle.Asn1.Crmf;

namespace KmsReportClient.Report.Basic
{
    public class MonitoringVCRProcessor : AbstractReportProcessor<ReportMonitoringVCR>
    {
        StackedHeaderDecorator DgvRender;
        string[] _notSaveCells = new string[] { "1","2.1", "2.2" };

        Dictionary<string, DataGridViewRow> _rows;

        MonitoringVCRPgDataDto[] _MonitoringVCRPGDataResult;

        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

         private string[] _columns = new string[]
        {
            "№ п/п;1",
            "Показатель;2",
            "Медико-экономическая экспертиза и экспертиза качества медицинской помощи;Экспертиза (обращения граждан на доступность и качество медицинской помощи)       ;3",
            "Медико-экономическая экспертиза и экспертиза качества медицинской помощи;Экспертиза (кроме обращений граждан на доступность и качество медицинcкой помощи)     ;4",
            "Медико-экономическая экспертиза и экспертиза качества медицинской помощи;Всего;5"
        };

        public MonitoringVCRProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
        base(inClient, dgv, cmb, txtb, page,
            XmlFormTemplate.MVCR.GetDescription(),
            Log,
            ReportGlobalConst.MonitoringVCR,
            reportsDictionary)
        {
            DgvRender = new StackedHeaderDecorator(Dgv);
            _rows = new Dictionary<string, DataGridViewRow>();
            InitReport();
        }

        private void FillTableByPgData()
        {
            try
            {
                var response = Client.GetMonitoringVCRPGData(new GetMonitoringVCRPGDataRequest
                {
                    Body = new GetMonitoringVCRPGDataRequestBody
                    {
                        fillial = FilialCode,
                        yymm = Report.Yymm
                    }
                });

                if (response != null)
                {

                    _MonitoringVCRPGDataResult = response.Body.GetMonitoringVCRPGDataResult;

                    for (int i = 0; i < Dgv.Rows.Count; i++)
                    {
                        var rowNum = Dgv.Rows[i].Cells[0].Value.ToString();
                        var pgRow = response.Body.GetMonitoringVCRPGDataResult.FirstOrDefault(x => x.RowNum == rowNum);
                        if (pgRow != null)
                        {
                            Dgv.Rows[i].Cells[2].Value = pgRow.ExpertWithEducation;
                            Dgv.Rows[i].Cells[3].Value = pgRow.ExpertWithoutEducation;
                            Dgv.Rows[i].Cells[4].Value = pgRow.Total;
                        }

                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public override AbstractReport CollectReportFromWs(string yymm)
        {
            var request = new GetReportRequest
            {
                Body = new GetReportRequestBody
                {
                    filialCode = FilialCode,
                    yymm = yymm,
                    reportType = ReportType.MVCR
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as ReportMonitoringVCR;
        }

        public override void FillDataGridView(string form)
        {
            if (Report != null)
            {
                if (Report.Data != null)
                {
                    foreach (DataGridViewRow row in Dgv.Rows)
                    {
                        var rowData = Report.Data.FirstOrDefault(x => x.RowNum == row.Cells[0].Value.ToString());
                        if (rowData != null)
                        {
                            row.Cells[2].Value = rowData.ExpertWithEducation;
                            row.Cells[3].Value = rowData.ExpertWithoutEducation;
                        }
                    }
                }
            }

            SetFormula();
            ValidReport();

        }
        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource)
        {

        }
        public override void InitReport()
        {
            Report = new ReportMonitoringVCR { Data = new MonitoringVCRData[Dgv.Rows.Count], IdType = IdReportType };
            for (int i = 0; i < Dgv.Rows.Count; i++)
            {
                Report.Data[i] = new MonitoringVCRData();
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
            var inReport = report as ReportMonitoringVCR;

            Report.IdReportData = inReport.IdReportData;
            Report.Data = inReport.Data;
        }

        public override void SaveReportDataSourceExcel()
        { }
        public override void SaveReportDataSourceHandle()
        { }

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
                    reportType = ReportType.MVCR
                }
            };


            var response = Client.SaveReport(request).Body.SaveReportResult as ReportMonitoringVCR;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;


        }
        public override void ToExcel(string filename, string filialName)
        {
            var excel = new ExcelMVCRCreator(filename, ExcelForm.MVCR, Report.Yymm, filialName, _rows);
            excel.CreateReport(Report, null);
        }
        public override string ValidReport()
        {
            string message = "";
            decimal control1 = 0,control2 = 0,control3 = 0;
            control1 = _rows.Where(x => x.Key == "2").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
            control2 = _rows.Where(x => x.Key == "2.1").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
            control3 = _rows.Where(x => x.Key == "2.2").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
            if (control1 < control2 + control3)
            { message += "Общее количество нарушений в строке '2' не может быть меньше суммы нарушений по строкам '2.1' и '2.2'"; }
            return message;
            
        }
        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            foreach (var clmn in _columns)
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

            foreach (var row in table)
            {
                var dgvRow = new DataGridViewRow();
                var N = new DataGridViewTextBoxCell { Value = row.RowNum_fromxml };
                var cellname = new DataGridViewTextBoxCell { Value = row.RowText_fromxml };
                dgvRow.Cells.Add(N);
                dgvRow.Cells.Add(cellname);
                int rowIndex = Dgv.Rows.Add(dgvRow);
            }

            FillTableByPgData();
            SetStyle();

            _rows = new Dictionary<string, DataGridViewRow>();
            foreach (DataGridViewRow row in Dgv.Rows)
            {
                _rows.Add(row.Cells[0].Value.ToString(), row);
            }
        }

        private void SetStyle()
        {

            Dgv.Columns[0].Width = 70;
            Dgv.Columns[1].Width = 150;
            Dgv.Columns[2].Width = 200;
            Dgv.Columns[3].Width = 200;
            Dgv.Columns[4].Width = 110;

            foreach (DataGridViewRow row in Dgv.Rows)
            {
                string rowNum = row.Cells[0].Value.ToString();
                if (_notSaveCells.Contains(rowNum))
                {
                    row.DefaultCellStyle.BackColor = Color.LightGray;
                    row.Cells[4].Style.BackColor = Color.DarkGray;
                    row.ReadOnly = true;
                    row.DefaultCellStyle.Font = new Font(Dgv.DefaultCellStyle.Font, FontStyle.Bold);
                }


            }
            Dgv.Columns[0].ReadOnly = Dgv.Columns[1].ReadOnly = true;
            Dgv.Columns[4].ReadOnly = Dgv.Columns[1].ReadOnly = true;
            Dgv.Columns[4].DefaultCellStyle.BackColor = Color.DarkGray;
        }

        protected override void FillReport(string form)
        {
            List<MonitoringVCRData> dataList = new List<MonitoringVCRData>();
            foreach (DataGridViewRow row in Dgv.Rows)
            {
                string rowNum = row.Cells[0].Value.ToString();
                if (_notSaveCells.Contains(rowNum))
                    continue;

                dataList.Add(new MonitoringVCRData
                {
                    RowNum = row.Cells[0].Value.ToString(),
                    ExpertWithEducation = GlobalUtils.TryParseDecimal(row.Cells[2].Value),
                    ExpertWithoutEducation = GlobalUtils.TryParseDecimal(row.Cells[3].Value),
                    Total = GlobalUtils.TryParseDecimal(row.Cells[4].Value),
                });

            }

            Report.Data = dataList.ToArray();
            SetFormula();
        }

        public void SetFormula()
        {
            foreach (var row in _rows.Reverse())
            {
                if (row.Key == "1")
                {
                    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "1.1" || x.Key == "1.2").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "1.1" || x.Key == "1.2").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                    row.Value.Cells[4].Value = _rows.Where(x => x.Key == "1.1" || x.Key == "1.2").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[4].Value));
                    continue;
                }


                if (row.Key == "2.1")
                {
                    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.1.1" || x.Key == "2.1.2" || x.Key == "2.1.3" || x.Key == "2.1.4" || x.Key == "2.1.5" || x.Key == "2.1.6" || x.Key == "2.1.7" || x.Key == "2.1.8" || x.Key == "2.1.9" || x.Key == "2.1.10").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                    //row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.1.1" || x.Key == "2.1.2" || x.Key == "2.1.3" || x.Key == "2.1.4" || x.Key == "2.1.5" || x.Key == "2.1.6" || x.Key == "2.1.7" || x.Key == "2.1.8" || x.Key == "2.1.9" || x.Key == "2.1.10").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                    row.Value.Cells[4].Value = _rows.Where(x => x.Key == "2.1.1" || x.Key == "2.1.2" || x.Key == "2.1.3" || x.Key == "2.1.4" || x.Key == "2.1.5" || x.Key == "2.1.6" || x.Key == "2.1.7" || x.Key == "2.1.8" || x.Key == "2.1.9" || x.Key == "2.1.10").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[4].Value));
                    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.1").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[4].Value) - GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                    continue;

                }


                if (row.Key == "2.2")
                {
                    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.2.1" || x.Key == "2.2.2" || x.Key == "2.2.3" || x.Key == "2.2.4" || x.Key == "2.2.5" || x.Key == "2.2.6" || x.Key == "2.2.7" || x.Key == "2.2.8").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                    //row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.2.1" || x.Key == "2.2.2" || x.Key == "2.2.3" || x.Key == "2.2.4" || x.Key == "2.2.5" || x.Key == "2.2.6" || x.Key == "2.2.7" || x.Key == "2.2.8").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                    row.Value.Cells[4].Value = _rows.Where(x => x.Key == "2.2.1" || x.Key == "2.2.2" || x.Key == "2.2.3" || x.Key == "2.2.4" || x.Key == "2.2.5" || x.Key == "2.2.6" || x.Key == "2.2.7" || x.Key == "2.2.8").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[4].Value));
                    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.2").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[4].Value)- GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                    continue;

                }




                if (_MonitoringVCRPGDataResult != null)
                {
                    MonitoringVCRPgDataDto dto = _MonitoringVCRPGDataResult.FirstOrDefault(x => x.RowNum == row.Key);
                    if (dto != null)
                    {
                         row.Value.Cells[3].Value = GlobalUtils.TryParseDecimal(row.Value.Cells[4].Value) - GlobalUtils.TryParseDecimal(row.Value.Cells[2].Value);
                    }
                }

                if (row.Key == "2")
                {
                    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[4].Value) - GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                }
            }
        }
    }
}
