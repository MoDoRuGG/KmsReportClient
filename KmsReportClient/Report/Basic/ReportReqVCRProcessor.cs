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
    public class ReportReqVCRProcessor : AbstractReportProcessor<ReportReqVCR>
    {

        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        private readonly string[] _columns = {"Наименование","№ строки","2019 г.","2020 г.","2021 г.","2022 г."," 7 месяцев 2023 г."};

        public ReportReqVCRProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
        base(inClient, dgv, cmb, txtb, page,
            XmlFormTemplate.ReqVCR.GetDescription(),
            Log,
            ReportGlobalConst.ReportReqVCR,
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
                    reportType = ReportType.ReqVCR
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as ReportReqVCR;

        }

        public override void FillDataGridView(string form)
        {
            FillDgwThemes(Dgv, form);
        }

        private void FillDgwThemes(DataGridView dgvReport, string form)
        {
            var reportReqVCRDto = Report.ReportDataList.SingleOrDefault(x => x.Theme.ToLower() == form.ToLower());

            if (reportReqVCRDto == null)
            {
                return;
            }

            if (reportReqVCRDto.Data == null || reportReqVCRDto.Data.Length == 0)
            {
                return;
            }

            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                var RowNum = row.Cells[1].Value.ToString();
                var data = reportReqVCRDto.Data.Single(x => x.RowNum == RowNum);
                if (data != null)
                {
                    row.Cells[2].Value = data.y2019;
                    row.Cells[3].Value = data.y2020;
                    row.Cells[4].Value = data.y2021;
                    row.Cells[5].Value = data.y2022;
                    row.Cells[6].Value = data.y2023;
                }
            }
        }

        public override void SaveReportDataSourceExcel()
        { }
        public override void SaveReportDataSourceHandle()
        { }

        public void SetFormula()
        {
        }


        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource)
        {

        }
        public override void InitReport()
        {
            Report = new ReportReqVCR { ReportDataList = new ReportReqVCRDto[ThemesList.Count], IdType = IdReportType };

            int i = 0;
            foreach (var theme in ThemesList.Select(x => x.Key))
            {
                var themeData = ThemeTextData.Tables_fromxml.Single(x => x.TableName_fromxml == theme);
                var rows = themeData.Rows_fromxml.Select(x => new ReportReqVCRDataDto { RowNum = x.RowNum_fromxml }).ToArray();

                Report.ReportDataList[i++] = new ReportReqVCRDto { Theme = theme, Data = rows };
            }
        }
        public override bool IsVisibleBtnDownloadExcel() => false;

        public override bool IsVisibleBtnHandle() => false;

        public override void MapForAutoFill(AbstractReport report)
        {
            if (report == null)
            {
                return;
            }
            var inReport = report as ReportReqVCR;

            var index = Report.ReportDataList.ToList().FindIndex(x => x.Theme == Cmb.Text);
            var inTheme = inReport.ReportDataList.Single(x => x.Theme == Cmb.Text);
            Report.ReportDataList[index] = inTheme;

        }
        public override void SaveToDb()
        {
            SetFormula();
            var request = new SaveReportRequest
            {
                Body = new SaveReportRequestBody
                {
                    filialCode = CurrentUser.FilialCode,
                    idUser = CurrentUser.IdUser,
                    report = Report,
                    yymm = Report.Yymm,
                    reportType = ReportType.ReqVCR
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportReqVCR;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
        }
        public override void ToExcel(string filename, string filialName)
        {

                //var excel = new ExcelReqVCRCreator(filename, ExcelForm.cadre, Report.Yymm, filialName, Client, FilialCode);
                //excel.CreateReport(Report, null);
        }
        public override string ValidReport() { return ""; }
        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            foreach (var clmn in _columns)
            {
                if (clmn == "Наименование")
                {
                    var column = new DataGridViewTextBoxColumn
                    {
                        HeaderText = clmn,
                        Width = 300,
                        DataPropertyName = "Indicator",
                        Name = "Indicator",
                        SortMode = DataGridViewColumnSortMode.NotSortable,
                        DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Azure }
                    };
                    Dgv.Columns.Add(column);
                }
                else
                {

                    var column = new DataGridViewTextBoxColumn
                    {
                        HeaderText = clmn,
                        Width = 60,
                        DataPropertyName = "Indicator",
                        Name = "Indicator",
                        SortMode = DataGridViewColumnSortMode.NotSortable,
                        DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.White }
                    };
                    Dgv.Columns.Add(column);
                }

                
            }

            int countRows = ThemeTextData.Tables_fromxml.Single(x => x.TableName_fromxml == form).RowsCount_fromxml;
            foreach (var row in table)
            {
                var dgvRow = new DataGridViewRow();
                var cellName = new DataGridViewTextBoxCell { Value = row.RowText_fromxml };
                var cellNum = new DataGridViewTextBoxCell { Value = row.RowNum_fromxml };
                dgvRow.Cells.Add(cellName);
                dgvRow.Cells.Add(cellNum);

                var exclusionCells = row.ExclusionCells_fromxml?.Split(',');
                for (int i = 2; i < countRows; i++)
                {
                    bool isNeedExcludeSum = exclusionCells?.Contains(i.ToString()) ?? false;
                    var cell = new DataGridViewTextBoxCell { Value = isNeedExcludeSum ? "x" : "0" };
                    dgvRow.Cells.Add(cell);

                    if (isNeedExcludeSum)
                    {
                        cell.ReadOnly = true;
                        cell.Style.BackColor = Color.DarkGray;
                    }
                }


                Dgv.Rows.Add(dgvRow);
            }

        }
        protected override void FillReport(string form)
        {
            if (form == null)
            {
                return;
            }
            FillThemesReport(Dgv, form);

        }

        private void FillThemesReport(DataGridView dgvReport, string form)
        {
            var reportReqVCRDto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportReqVCRDto != null)
            {

                var reportData = new List<ReportReqVCRDataDto>();
                foreach (DataGridViewRow row in dgvReport.Rows)
                {
                    var RowNum = row.Cells[1].Value.ToString();
                    var data = new ReportReqVCRDataDto
                    {
                        RowNum = RowNum,
                        y2019 = GlobalUtils.TryParseDecimal(row.Cells[2].Value),
                        y2020 = GlobalUtils.TryParseDecimal(row.Cells[3].Value),
                        y2021 = GlobalUtils.TryParseDecimal(row.Cells[4].Value),
                        y2022 = GlobalUtils.TryParseDecimal(row.Cells[5].Value),
                        y2023 = GlobalUtils.TryParseDecimal(row.Cells[6].Value),
                    };
                    reportData.Add(data);
                }

                reportReqVCRDto.Data = reportData.ToArray();
            }
        }
    }
}
