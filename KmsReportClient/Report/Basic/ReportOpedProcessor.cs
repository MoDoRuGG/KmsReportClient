using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
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
    public class ReportOpedProcessor : AbstractReportProcessor<ReportOped>
    {
        #region Приватные переменные и объекты
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        private string[] columns = new string[] { "АПП", "Стационар", "Стационарозамещая помощь", "Скорая медицинская помощь", "Примечания" };
        private ReportOpedDto[] firstValueYear = new ReportOpedDto[] { };

        List<CellModel> beforeJunyNormativ = new List<CellModel>()
        {
            new CellModel
            {
                Row = 1,
                Column =2,
                Value = 0.8M,

            },
            new CellModel
            {
                Row = 1,
                Column =3,
                Value = 8M,

            },

            new CellModel
            {
                Row = 1,
                Column =4,
                Value = 8M,

            },
               new CellModel
            {
                Row = 1,
                Column =5,
                Value = 3M,

            },


                new CellModel
            {
                Row = 2,
                Column =2,
                Value = 0.5M,

            },
            new CellModel
            {
                Row = 2,
                Column =3,
                Value = 5M,

            },

            new CellModel
            {
                Row = 2,
                Column =4,
                Value = 3M,

            },
             new CellModel
            {
                Row = 2,
                Column =5,
                Value = 1.5M,

            },


        };

        List<CellModel> afterJunyNormativ  = new List<CellModel>()
        {
            new CellModel
            {
                Row = 1,
                Column =2,
                Value = 0.5M,

            },
            new CellModel
            {
                Row = 1,
                Column =3,
                Value = 6M,

            },

            new CellModel
            {
                Row = 1,
                Column =4,
                Value = 6M,

            },
               new CellModel
            {
                Row = 1,
                Column =5,
                Value = 2M,

            },


                new CellModel
            {
                Row = 2,
                Column =2,
                Value = 0.2M,

            },
            new CellModel
            {
                Row = 2,
                Column =3,
                Value = 3M,

            },

            new CellModel
            {
                Row = 2,
                Column =4,
                Value = 1.5M,

            },
             new CellModel
            {
                Row = 2,
                Column =5,
                Value = 0.5M,

            },


        };



        private int[] rowWithFormula = new int[] { 6, 7, 8, 9 };


        #endregion

        public ReportOpedProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
           base(inClient, dgv, cmb, txtb, page,
               XmlFormTemplate.Oped.GetDescription(),
               Log,
               ReportGlobalConst.ReportOped,
               reportsDictionary)
        {
            InitReport();
            CreateCmbItem();

        }
        public override AbstractReport CollectReportFromWs(string yymm)
        {

            var request = new GetReportRequest
            {
                Body = new GetReportRequestBody
                {
                    filialCode = FilialCode,
                    yymm = yymm,
                    reportType = ReportType.Oped
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as ReportOped;
        }

        public override void FillDataGridView(string form)
        {
            if (form == null)
            {
                return;
            }

            if (Report.ReportDataList != null && Report.ReportDataList.Length > 0)
            {
                if (form == "Свод")
                {
                    var yearData = Client.GetYearOpedData(new GetYearOpedDataRequest
                    {
                        Body = new GetYearOpedDataRequestBody
                        {
                            yymm = Report.Yymm,
                            filiall = FilialCode
                        }
                    });

                    firstValueYear = yearData.Body.GetYearOpedDataResult;
                }

                foreach (DataGridViewRow row in Dgv.Rows)
                {
                    var rowNum = row.Cells[0].Value.ToString();
                    //Console.WriteLine(rowNum);

                    var data = Report.ReportDataList.SingleOrDefault(x => x.RowNum.ToString() == rowNum);
                    if (form == "Свод")
                    {
                        data = firstValueYear.SingleOrDefault(x => x.RowNum.ToString() == rowNum);
                    }

                    if (data != null)
                    {
                        row.Cells[2].Value = (int)data.App;
                        row.Cells[3].Value = (int)data.Ks;
                        row.Cells[4].Value = (int)data.Ds;
                        row.Cells[5].Value = (int)data.Smp;
                        row.Cells[6].Value = data.Notes;

                    }
                }

                SetStaticValue();
                SetFormula();
                SetTotalColumn();

            }
        }


        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status)
        {

        }

        public override void InitReport()
        {
            Report = new ReportOped { ReportDataList = new ReportOpedDto[] { }, IdType = IdReportType };

        }

        public override bool IsVisibleBtnDownloadExcel()
        {
            return false;
        }

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
                    reportType = ReportType.Oped
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportOped;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
        }

        public override void ToExcel(string filename, string filialName)
        {
            var excel = new ExcelOpedCreator(filename, ExcelForm.Oped, Report.Yymm, filialName);
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
            }

            Dgv.Rows[5].ReadOnly = true;
            Dgv.Rows[6].ReadOnly = true;
            Dgv.Rows[7].ReadOnly = true;
            Dgv.Rows[8].ReadOnly = true;

            Dgv.Rows[5].Cells[6].ReadOnly = false;
            Dgv.Rows[6].Cells[6].ReadOnly = false;
            Dgv.Rows[7].Cells[6].ReadOnly = false;
            Dgv.Rows[8].Cells[6].ReadOnly = false;

            SetStaticValue();
            SetStyleDgv();
        }

        protected override void FillReport(string form)
        {
            int[] exCells = new int[] { 6, 7, 8, 9 };
            if (form == null || form == "Свод")
            {
                return;
            }

            var reportDto = new List<ReportOpedDto>();

            foreach (DataGridViewRow row in Dgv.Rows)
            {
                int rowNum = Convert.ToInt32(row.Cells[0].Value);

                if (!exCells.Contains(rowNum))
                {

                    var data = new ReportOpedDto
                    {
                        RowNum = row.Cells[0].Value.ToString(),
                        App = GlobalUtils.TryParseInt(row.Cells[2].Value),
                        Ks = GlobalUtils.TryParseInt(row.Cells[3].Value),
                        Ds = GlobalUtils.TryParseInt(row.Cells[4].Value),
                        Smp = GlobalUtils.TryParseInt(row.Cells[5].Value),
                        Notes = row.Cells[6].Value?.ToString() ?? ""
                    };
                    reportDto.Add(data);
                }
                else
                {
                    var data = new ReportOpedDto
                    {
                        RowNum = row.Cells[0].Value.ToString(),
                        App = 0,
                        Ks = 0,
                        Ds = 0,
                        Smp = 0,
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


        private void SetStyleDgv()
        {
            Dgv.Rows[1].DefaultCellStyle.BackColor = Color.LightGray;
            Dgv.Rows[2].DefaultCellStyle.BackColor = Color.LightGray;

            Dgv.Rows[5].DefaultCellStyle.BackColor = Color.LightGray;
            Dgv.Rows[6].DefaultCellStyle.BackColor = Color.LightGray;
            Dgv.Rows[7].DefaultCellStyle.BackColor = Color.LightGray;
            Dgv.Rows[8].DefaultCellStyle.BackColor = Color.LightGray;

            Dgv.Rows[5].Cells[6].Style.BackColor = Color.White;
            Dgv.Rows[6].Cells[6].Style.BackColor = Color.White;
            Dgv.Rows[7].Cells[6].Style.BackColor = Color.White;
            Dgv.Rows[8].Cells[6].Style.BackColor = Color.White;



        }


        private void SetStaticValue()
        {

            List<CellModel> normativ;

            if (Convert.ToInt32(Report.Yymm) > 2105)
            {
                normativ = afterJunyNormativ;
            }
            else
            {
                normativ = beforeJunyNormativ;

            }


            foreach (var data in normativ)
            {
                Dgv.Rows[data.Row].Cells[data.Column].Value = data.Value;
            }
        }

        private void SetFormula()
        {

            var row1 = Dgv.Rows[0];
            var row4 = Dgv.Rows[3];
            var row5 = Dgv.Rows[4];



            foreach (var row in rowWithFormula)
            {
                int rowDgv = row - 1;
                if (rowDgv == 5)
                {

                    Dgv.Rows[rowDgv].Cells[2].Value = GetValueFormula(row1.Cells[2], row4.Cells[2]);
                    Dgv.Rows[rowDgv].Cells[3].Value = GetValueFormula(row1.Cells[3], row4.Cells[3]);
                    Dgv.Rows[rowDgv].Cells[4].Value = GetValueFormula(row1.Cells[4], row4.Cells[4]);
                    Dgv.Rows[rowDgv].Cells[5].Value = GetValueFormula(row1.Cells[5], row4.Cells[5]);

                }

                if (rowDgv == 6)
                {
                    Dgv.Rows[rowDgv].Cells[2].Value = GetValueFormula2(Dgv.Rows[5].Cells[2].Value, Dgv.Rows[1].Cells[2].Value);
                    Dgv.Rows[rowDgv].Cells[3].Value = GetValueFormula2(Dgv.Rows[5].Cells[3].Value, Dgv.Rows[1].Cells[3].Value);
                    Dgv.Rows[rowDgv].Cells[4].Value = GetValueFormula2(Dgv.Rows[5].Cells[4].Value, Dgv.Rows[1].Cells[4].Value);
                    Dgv.Rows[rowDgv].Cells[5].Value = GetValueFormula2(Dgv.Rows[5].Cells[5].Value, Dgv.Rows[1].Cells[5].Value);
                }


                if (rowDgv == 7)
                {

                    Dgv.Rows[rowDgv].Cells[2].Value = GetValueFormula(row1.Cells[2], row5.Cells[2]);
                    Dgv.Rows[rowDgv].Cells[3].Value = GetValueFormula(row1.Cells[3], row5.Cells[3]);
                    Dgv.Rows[rowDgv].Cells[4].Value = GetValueFormula(row1.Cells[4], row5.Cells[4]);
                    Dgv.Rows[rowDgv].Cells[5].Value = GetValueFormula(row1.Cells[5], row5.Cells[5]);
                }

                if (rowDgv == 8)
                {

                    Dgv.Rows[rowDgv].Cells[2].Value = GetValueFormula2(Dgv.Rows[7].Cells[2].Value, Dgv.Rows[2].Cells[2].Value);
                    Dgv.Rows[rowDgv].Cells[3].Value = GetValueFormula2(Dgv.Rows[7].Cells[3].Value, Dgv.Rows[2].Cells[3].Value);
                    Dgv.Rows[rowDgv].Cells[4].Value = GetValueFormula2(Dgv.Rows[7].Cells[4].Value, Dgv.Rows[2].Cells[4].Value);
                    Dgv.Rows[rowDgv].Cells[5].Value = GetValueFormula2(Dgv.Rows[7].Cells[5].Value, Dgv.Rows[2].Cells[5].Value);
                }


            }

        }
        public string GetValueFormula(DataGridViewCell value1, DataGridViewCell value2)
        {

            string result = "";

            try
            {
                decimal val1 = value1.Value == null ? 0 : Convert.ToDecimal(value1.Value);
                decimal val2 = value2.Value == null ? 0 : Convert.ToDecimal(value2.Value);



                if (val1 != 0)
                {
                    result = Math.Round((val2 * 100) / val1, 2).ToString();
                }
                else
                {
                    result = "Деление на 0";

                }

                return result;

            }
            catch (Exception ex)
            {
                return result;
            }



        }

        public string GetValueFormula2(object value1, object value2)
        {
            string result = "";


            try
            {

                if (Convert.ToDecimal(value1) == Convert.ToDecimal(value2))
                {
                    result = "Выполнен";
                    return result;

                }

                if (Convert.ToDecimal(value1) > Convert.ToDecimal(value2))
                {
                    result = "Выполнен";
                }
                else
                {
                    result = " Не выполнен";
                }
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                result = "Ошибка!";
                return result;
            }


        }

        public override void CalculateCells()
        {
            SetFormula();

        }

        public void CreateCmbItem()
        {

        }
    }
}

