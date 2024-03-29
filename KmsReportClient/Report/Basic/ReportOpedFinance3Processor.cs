﻿using System;
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
    public class ReportOpedFinance3Processor : AbstractReportProcessor<ReportOpedFinance3>
    {

        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        private readonly string[] _columns = { "Наименование показателя", "№ строки", "Примечание" };
        public ReportOpedFinance3Processor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
          base(inClient, dgv, cmb, txtb, page,
              XmlFormTemplate.OpedFinance3.GetDescription(),
              Log,
              ReportGlobalConst.ReportOpedFinance3,
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
                    reportType = ReportType.OpedFinance3
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as ReportOpedFinance3;
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
                        row.Cells[2].Value = data.Notes;
                    }
                    else
                    {
                        row.Cells[2].Value = "";
                    }
                }

                //CalculateCells();

            }
        }

        public override void SaveReportDataSourceExcel()
        { }
        public override void SaveReportDataSourceHandle()
        { }
        //public void CalculateCells()
        //{
        //        var row1 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "1.");
        //}

        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource)
        {

        }
        public override void InitReport()
        {
            Report = new ReportOpedFinance3 { ReportDataList = Array.Empty<ReportOpedFinance3Data>(), IdType = IdReportType };
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
                    reportType = ReportType.OpedFinance3
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportOpedFinance3;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
        }
        public override void ToExcel(string filename, string filialName)
        {
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

            var row1 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "1.");

            row1.Cells[1].Style.BackColor = Color.LightGray;
            row1.Cells[1].ReadOnly = true;

        }
        protected override void FillReport(string form)
        {

            if (form == null)
            {
                return;
            }

            var reportDto = new List<ReportOpedFinance3Data>();

            foreach (DataGridViewRow row in Dgv.Rows)
            {
                var data = new ReportOpedFinance3Data
                {
                    RowNum = row.Cells[1].Value.ToString(),
                    Notes = row.Cells[2].Value.ToString()

                };
                reportDto.Add(data);
            }

            Report.ReportDataList = reportDto.ToArray();
        }
    }
}
