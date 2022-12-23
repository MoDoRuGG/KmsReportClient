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
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.XML;
using KmsReportClient.Support;
using NLog;

namespace KmsReportClient.Report.Basic
{
    public class FSSMonitoringProcessor : AbstractReportProcessor<ReportFSSMonitroing>
    {
        StackedHeaderDecorator DgvRender;
        string[] _notSaveCells = new string[] { "1", "1.1", "1.2", "2", "2.1", "2.2" };

        Dictionary<string, DataGridViewRow> _rows;

        FSSMonitoringPgDataDto[] _FSSMonitoringPGDataResult;

        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

         private string[] _columns = new string[]
        {
            "№ п/п",
            "Показатель",
            "Медико-экономическая экспертиза и экспертиза качества медицинской помощи;Экспертиза (обращения граждан на доступность и качество медицинской помощи)",
            "Медико-экономическая экспертиза и экспертиза качества медицинской помощи;Экспертиза (кроме обращений граждан на доступность и качество медицинкой помощи)",
            "Медико-экономическая экспертиза и экспертиза качества медицинской помощи;Всего"
        };

        public FSSMonitoringProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
        base(inClient, dgv, cmb, txtb, page,
            XmlFormTemplate.MFSS.GetDescription(),
            Log,
            ReportGlobalConst.FSSMonitoring,
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
                var response = Client.GetFSSMonitoringPGData(new GetFSSMonitoringPGDataRequest
                {
                    Body = new GetFSSMonitoringPGDataRequestBody
                    {
                        fillial = FilialCode,
                        yymm = Report.Yymm
                    }
                });

                if (response != null)
                {

                    _FSSMonitoringPGDataResult = response.Body.GetFSSMonitoringPGDataResult;

                    for (int i = 0; i < Dgv.Rows.Count; i++)
                    {
                        var rowNum = Dgv.Rows[i].Cells[0].Value.ToString();
                        var pgRow = response.Body.GetFSSMonitoringPGDataResult.FirstOrDefault(x => x.RowNum == rowNum);
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
                    reportType = ReportType.MFSS
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as ReportFSSMonitroing;
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

        }
        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource)
        {

        }
        public override void InitReport()
        {
            Report = new ReportFSSMonitroing { Data = new FSSMonitroingData[Dgv.Rows.Count], IdType = IdReportType };
            for (int i = 0; i < Dgv.Rows.Count; i++)
            {
                Report.Data[i] = new FSSMonitroingData();
            }
        }
        public override bool IsVisibleBtnDownloadExcel() => true;

        public override bool IsVisibleBtnHandle() => false;


        public override void MapForAutoFill(AbstractReport report)
        {
            if (report == null)
            {
                return;
            }
            var inReport = report as ReportFSSMonitroing;

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
                    reportType = ReportType.MFSS
                }
            };


            var response = Client.SaveReport(request).Body.SaveReportResult as ReportFSSMonitroing;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;


        }
        public override void ToExcel(string filename, string filialName)
        {
            var excel = new ExcelMFSSCreator(filename, ExcelForm.MFSS, Report.Yymm, filialName, _rows);
            excel.CreateReport(Report, null);
        }
        public override string ValidReport()
        {
            return "";
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
                var N = new DataGridViewTextBoxCell { Value = row.Num };
                var cellname = new DataGridViewTextBoxCell { Value = row.Name };
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

                if(rowNum == "2.1.8ИНФ" || rowNum == "2.2.4ИНФ")
                {
                    row.ReadOnly = true;
                }

            }
            Dgv.Columns[4].DefaultCellStyle.BackColor = Color.DarkGray;
            Dgv.Columns[4].ReadOnly = Dgv.Columns[0].ReadOnly = Dgv.Columns[1].ReadOnly = true;
        }

        protected override void FillReport(string form)
        {
            List<FSSMonitroingData> dataList = new List<FSSMonitroingData>();
            foreach (DataGridViewRow row in Dgv.Rows)
            {
                string rowNum = row.Cells[0].Value.ToString();
                if (_notSaveCells.Contains(rowNum))
                    continue;

                dataList.Add(new FSSMonitroingData
                { 
                    RowNum = row.Cells[0].Value.ToString(),
                    ExpertWithEducation = GlobalUtils.TryParseDecimal(row.Cells[2].Value),
                    ExpertWithoutEducation = GlobalUtils.TryParseDecimal(row.Cells[3].Value),
                    Total = GlobalUtils.TryParseDecimal(row.Cells[4].Value),
                });

            }

            Report.Data = dataList.ToArray();
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

                if (row.Key == "2")
                {
                    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.1" || x.Key == "2.2").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.1" || x.Key == "2.2").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                  

                }


                if (row.Key == "2.1")
                {
                    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.1.1" || x.Key == "2.1.2" || x.Key == "2.1.3" || x.Key == "2.1.4" || x.Key == "2.1.5" || x.Key == "2.1.6" || x.Key == "2.1.7" || x.Key == "2.1.8").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.1.1" || x.Key == "2.1.2" || x.Key == "2.1.3" || x.Key == "2.1.4" || x.Key == "2.1.5" || x.Key == "2.1.6" || x.Key == "2.1.7" || x.Key == "2.1.8").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                    row.Value.Cells[4].Value = _rows.Where(x => x.Key == "2.1.1" || x.Key == "2.1.2" || x.Key == "2.1.3" || x.Key == "2.1.4" || x.Key == "2.1.5" || x.Key == "2.1.6" || x.Key == "2.1.7" || x.Key == "2.1.8").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[4].Value));
                    continue;

                }


                if (row.Key == "2.2")
                {
                    row.Value.Cells[2].Value = _rows.Where(x => x.Key == "2.2.1" || x.Key == "2.2.2" || x.Key == "2.2.3" || x.Key == "2.2.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                    row.Value.Cells[3].Value = _rows.Where(x => x.Key == "2.2.1" || x.Key == "2.2.2" || x.Key == "2.2.3" || x.Key == "2.2.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                    row.Value.Cells[4].Value = _rows.Where(x => x.Key == "2.2.1" || x.Key == "2.2.2" || x.Key == "2.2.3" || x.Key == "2.2.4").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[4].Value));
                    continue;

                }


                if (_FSSMonitoringPGDataResult != null)
                {

                    // ПО ЗАПРОСУ ГУЖЕНКО перевожу все на суммирование, без подтягивания данных ПГ
                    
                    FSSMonitoringPgDataDto dto = _FSSMonitoringPGDataResult.FirstOrDefault(x => x.RowNum == row.Key);
                    if (row.Key != "2.2.4ИНФ" || row.Key != "2.1.8ИНФ") { row.Value.Cells[4].Value = GlobalUtils.TryParseDecimal(row.Value.Cells[2].Value) + GlobalUtils.TryParseDecimal(row.Value.Cells[3].Value); }
                    // if (dto != null)
                    else
                    {
                        if (GlobalUtils.TryParseDecimal(dto.Total) == 0.00m) // Если по ПГ нам ничего не пришло, то можно суммировать
                        {
                            
                        }
                        else
                        {
                            if (GlobalUtils.TryParseDecimal(row.Value.Cells[2].Value) == 0.00m && GlobalUtils.TryParseDecimal(row.Value.Cells[3].Value) != 0.00m)
                            {
                                row.Value.Cells[2].Value = GlobalUtils.TryParseDecimal(row.Value.Cells[4].Value) - GlobalUtils.TryParseDecimal(row.Value.Cells[3].Value);
                            }
                            else if (GlobalUtils.TryParseDecimal(row.Value.Cells[3].Value) == 0.00m && GlobalUtils.TryParseDecimal(row.Value.Cells[2].Value) != 0.00m)
                            {
                                row.Value.Cells[3].Value = GlobalUtils.TryParseDecimal(row.Value.Cells[4].Value) - GlobalUtils.TryParseDecimal(row.Value.Cells[2].Value);
                            }
                            //else if (GlobalUtils.TryParseDecimal(row.Value.Cells[2].Value) != 0.00m && GlobalUtils.TryParseDecimal(row.Value.Cells[3].Value) != 0.00m && (GlobalUtils.TryParseDecimal(row.Value.Cells[2].Value) + GlobalUtils.TryParseDecimal(row.Value.Cells[3].Value) != GlobalUtils.TryParseDecimal(row.Value.Cells[4].Value)))
                            //{
                            //    row.Value.Cells[4].Value = GlobalUtils.TryParseDecimal(row.Value.Cells[2].Value) + GlobalUtils.TryParseDecimal(row.Value.Cells[3].Value);
                            //}


                        }

                    }
                }

            }
        }
    }
}
