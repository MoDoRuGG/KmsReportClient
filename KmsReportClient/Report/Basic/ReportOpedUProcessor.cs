using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
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

namespace KmsReportClient.Report.Basic
{
    public class ReportOpedUProcessor : AbstractReportProcessor<ReportOpedU>
    {
        #region Приватные переменные и объекты
        int[] _notSaveRow = { 0, 4, 8 };

        int[] _calcRows = { 3, 7, 11 };
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        private string[] columns = new string[] {
            "АПП",
            "Стационар",
            "Стационарозамещающая помощь",
            "Скорая медицинская помощь",
            "Примечания" };
        //private ReportOpedUDto[] firstValueYear = new ReportOpedUDto[] { };

        //List<CellModel> beforeJunyNormativ = new List<CellModel>()
        //{
        //    new CellModel
        //    {
        //        Row = 1,
        //        Column =2,
        //        Value = 0.8M,

        //    },
        //    new CellModel
        //    {
        //        Row = 1,
        //        Column =3,
        //        Value = 8M,

        //    },

        //    new CellModel
        //    {
        //        Row = 1,
        //        Column =4,
        //        Value = 8M,

        //    },
        //       new CellModel
        //    {
        //        Row = 1,
        //        Column =5,
        //        Value = 3M,

        //    },


        //        new CellModel
        //    {
        //        Row = 2,
        //        Column =2,
        //        Value = 0.5M,

        //    },
        //    new CellModel
        //    {
        //        Row = 2,
        //        Column =3,
        //        Value = 5M,

        //    },

        //    new CellModel
        //    {
        //        Row = 2,
        //        Column =4,
        //        Value = 3M,

        //    },
        //     new CellModel
        //    {
        //        Row = 2,
        //        Column =5,
        //        Value = 1.5M,

        //    },


        //};

        //List<CellModel> afterJunyNormativ = new List<CellModel>()
        //{
        //    new CellModel
        //    {
        //        Row = 1,
        //        Column =2,
        //        Value = 0.5M,

        //    },
        //    new CellModel
        //    {
        //        Row = 1,
        //        Column =3,
        //        Value = 6M,

        //    },

        //    new CellModel
        //    {
        //        Row = 1,
        //        Column =4,
        //        Value = 6M,

        //    },
        //       new CellModel
        //    {
        //        Row = 1,
        //        Column =5,
        //        Value = 2M,

        //    },


        //        new CellModel
        //    {
        //        Row = 2,
        //        Column =2,
        //        Value = 0.2M,

        //    },
        //    new CellModel
        //    {
        //        Row = 2,
        //        Column =3,
        //        Value = 3M,

        //    },

        //    new CellModel
        //    {
        //        Row = 2,
        //        Column =4,
        //        Value = 1.5M,

        //    },
        //     new CellModel
        //    {
        //        Row = 2,
        //        Column =5,
        //        Value = 0.5M,

        //    },


        //};



        //private int[] rowWithFormula = new int[] { 6, 7, 8, 9 };


        #endregion

        public ReportOpedUProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
           base(inClient, dgv, cmb, txtb, page,
               XmlFormTemplate.OpedU.GetDescription(),
               Log,
               ReportGlobalConst.ReportOpedU,
               reportsDictionary)
        {
            //DgvRender = new StackedHeaderDecorator(Dgv);
            InitReport();
            //CreateCmbItem();

        }
        public override AbstractReport CollectReportFromWs(string yymm)
        {

            var request = new GetReportRequest
            {
                Body = new GetReportRequestBody
                {
                    filialCode = FilialCode,
                    yymm = yymm,
                    reportType = ReportType.OpedU
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as ReportOpedU;
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
                    var rowNum = row.Cells[0].Value.ToString();
                    //Console.WriteLine(rowNum);

                    var data = Report.ReportDataList.SingleOrDefault(x => x.RowNum.ToString() == rowNum);
                    //if (form == "Свод")
                    //{
                    //    data = firstValueYear.SingleOrDefault(x => x.RowNum.ToString() == rowNum);
                    //}

                    if (data != null)
                    {
                        row.Cells[2].Value = (int)data.App;
                        row.Cells[3].Value = (int)data.Ks;
                        row.Cells[4].Value = (int)data.Ds;
                        row.Cells[5].Value = (int)data.Smp;
                        //row.Cells[6].Value = (int)data.AppOnco;
                        //row.Cells[7].Value = (int)data.KsOnco;
                        //row.Cells[8].Value = (int)data.DsOnco;
                        //row.Cells[9].Value = (int)data.SmpOnco;
                        //row.Cells[10].Value = (int)data.AppLeth;
                        //row.Cells[11].Value = (int)data.KsLeth;
                        //row.Cells[12].Value = (int)data.DsLeth;
                        //row.Cells[13].Value = (int)data.SmpLeth;
                        row.Cells[6].Value = data.Notes;

                    }
                }

                SetCalculateValue();
            }
        }

        public void SetCalculateValue()
        {
            foreach (int row in _calcRows)
            {
                for (int i = 2; i < Dgv.Rows[row].Cells.Count - 1; i++)
                {
                    try
                    {
                        Dgv.Rows[row].Cells[i].Value = Math.Round(GlobalUtils.TryParseDecimal(Dgv.Rows[row - 1].Cells[i].Value) / GlobalUtils.TryParseDecimal(Dgv.Rows[row - 2].Cells[i].Value) * 100, 2);

                    }
                    catch (Exception) { }
                }
            }
        }
        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource)
        {

        }
        public override void SaveReportDataSourceExcel()
        { }
        public override void SaveReportDataSourceHandle()
        { }
        public override void InitReport()
        {
            Report = new ReportOpedU { ReportDataList = new ReportOpedUDto[] { }, IdType = IdReportType };
        }

        public override bool IsVisibleBtnDownloadExcel() => false;

        public override bool IsVisibleBtnHandle() => false;

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
                    reportType = ReportType.OpedU
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportOpedU;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
        }

        public override void ToExcel(string filename, string filialName)
        {
            var excel = new ExcelOpedUCreator(filename, ExcelForm.OpedU, Report.Yymm, filialName);
            excel.CreateReport(Report, null);
        }

        public override string ValidReport()
        {
            return "";

        }

        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            CreateDgvCommonColumns(Dgv, 50);
            foreach (var row in table)
            {

                var dgvRow = new DataGridViewRow();
                var N = new DataGridViewTextBoxCell { Value = row.Num };
                var cellname = new DataGridViewTextBoxCell { Value = row.Name };
                dgvRow.Cells.Add(N);
                dgvRow.Cells.Add(cellname);
                int rowIndex = Dgv.Rows.Add(dgvRow);
            };




            //Dgv.Rows[5].ReadOnly = true;
            //Dgv.Rows[6].ReadOnly = true;
            //Dgv.Rows[7].ReadOnly = true;
            //Dgv.Rows[8].ReadOnly = true;

            //Dgv.Rows[5].Cells[6].ReadOnly = false;
            //Dgv.Rows[6].Cells[6].ReadOnly = false;
            //Dgv.Rows[7].Cells[6].ReadOnly = false;
            //Dgv.Rows[8].Cells[6].ReadOnly = false;

            //SetStaticValue();
            //SetStyleDgv();


            foreach (int row in _notSaveRow)
            {
                Dgv.Rows[row].DefaultCellStyle.BackColor = Color.LightGray;
                Dgv.Rows[row].ReadOnly = true;
            }

            foreach (int row in _calcRows)
            {
                Dgv.Rows[row].DefaultCellStyle.BackColor = Color.LightGreen;
                Dgv.Rows[row].ReadOnly = true;
            }
        }

        protected override void FillReport(string form)
        {
            int[] _notSaveRow = { 0, 4, 8 };

            if (form == null || form == "Свод")
            {
                return;
            }

            var reportDto = new List<ReportOpedUDto>();

            foreach (DataGridViewRow row in Dgv.Rows)
            {
                if (!_notSaveRow.Contains(row.Index) && !_calcRows.Contains(row.Index))
                {
                    var data = new ReportOpedUDto
                    {
                        RowNum = row.Cells[0].Value.ToString(),
                        App = GlobalUtils.TryParseInt(row.Cells[2].Value),
                        Ks = GlobalUtils.TryParseInt(row.Cells[3].Value),
                        Ds = GlobalUtils.TryParseInt(row.Cells[4].Value),
                        Smp = GlobalUtils.TryParseInt(row.Cells[5].Value),
                        //AppOnco = GlobalUtils.TryParseInt(row.Cells[6].Value),
                        //KsOnco = GlobalUtils.TryParseInt(row.Cells[7].Value),
                        //DsOnco = GlobalUtils.TryParseInt(row.Cells[8].Value),
                        //SmpOnco = GlobalUtils.TryParseInt(row.Cells[9].Value),
                        //AppLeth = GlobalUtils.TryParseInt(row.Cells[10].Value),
                        //KsLeth = GlobalUtils.TryParseInt(row.Cells[11].Value),
                        //DsLeth = GlobalUtils.TryParseInt(row.Cells[12].Value),
                        //SmpLeth = GlobalUtils.TryParseInt(row.Cells[13].Value),
                        Notes = row.Cells[6].Value?.ToString() ?? ""
                    };
                    reportDto.Add(data);
                }
            }

            Report.ReportDataList = reportDto.ToArray();

        }

        private void CreateDgvCommonColumns(DataGridView dgvReport, int widthFirstColumn)
        {
            dgvReport.AllowUserToAddRows = false;
            dgvReport.ColumnHeadersVisible = true;
            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = "№",
                Width = 40,
                DataPropertyName = "NumRow",
                Name = "NumRow",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Наименование показателя",
                Width = 350,
                DataPropertyName = "Indicator",
                Name = "Indicator",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);

            foreach (var col in columns)
            {
                var dgvColumn = new DataGridViewTextBoxColumn
                {
                    HeaderText = col,
                    Width = 150,
                    ReadOnly = false,
                    SortMode = DataGridViewColumnSortMode.NotSortable
                };
                dgvReport.Columns.Add(dgvColumn);
            }
        }

    }
}

