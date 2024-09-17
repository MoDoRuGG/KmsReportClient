using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
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
    class ReportDoffProcessor : AbstractReportProcessor<ReportDoff>
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        private readonly string[] _forms2346 = { "Таблица 2", "Таблица 3", "Таблица 4", "Таблица 6" };
        //private readonly string[] _forms3 = { "Таблица 3" };
        private readonly string[] _forms31_41 = { "Таблица 3.1", "Таблица 4.1" };
        //private readonly string[] _forms4 = { "Таблица 4" };
        //private readonly string[] _forms41 = { "Таблица 4.1" };
        //private readonly string[] _forms6 = { "Таблица 6" };


        private readonly string[][] _headers = {
            new[]
            { "Информация" }, //1
            new[]
            { "За отчетный период (месяц)", "Всего, с начала календарного года","Всего, с даты заключения соглашения" }, //2
            new[]
            { "За отчетный период (месяц)", "Всего, с начала календарного года","Всего, с даты заключения соглашения" }, //3
            new[]
            { "Количество обратившихся участников СВО", "Количество обратившихся членов семей" }, //3.1
            new[]
            { "За отчетный период (месяц)", "Всего, с начала календарного года","Всего, с даты заключения соглашения" }, //4
            new[]
            { "Количество обратившихся участников СВО", "Количество обратившихся членов семей" }, //4.1
            new[]
            { "Информация" }, //5
            new[]
            { "За отчетный период (месяц)", "Всего, с начала календарного года","Всего, с даты заключения соглашения" }, //6
            new[]
            { "Дата проведения (дд.мм.гггг)", "Вид мероприятия, наименование, инициатор (руководитель мероприятия)", "Краткое содержание, относящееся к тематике соглашения (участники, темы обсуждения, решения)"  }, //7
            new[]
            { "Куда направлено", "Реквизиты документа (при наличии)", "Краткое содержание (суть предложений)" }, //8
            new[]
            { "Информация" }, //9


        };

        private readonly Dictionary<string, string> _headersMap = new Dictionary<string, string>
        {
            { "Таблица 1", "Огранизация рабочего места страхового представителя в Филиале Фонда" },
            { "Таблица 2", "Количество обращений" },
            { "Таблица 3", "Количество обратившихся за ИИС (завершенные обращения)" },
            { "Таблица 3.1", "Тема (предмет) обращения" },
            { "Таблица 4", "Количество индивидуально проинформированных застрахованных лиц" },
            { "Таблица 4.1", "Повод информирования" },
            { "Таблица 5", "Сведения об издании и распространении информационных материалов, в том числе индивидуального характера, о правах в сфере охраны здоровья и ОМС: памятки, брошюры, листовки" },
            { "Таблица 6", "Количество завершённых рассмотрением обоснованных жалоб, по которым проведена ЭКМП" },
            { "Таблица 7", "Сведения о проведенных совместных с Фондом рабочих встречах и совещаниях, об участии в рабочих группах, семинарах, круглых столах и прочих мероприятиях" },
            { "Таблица 8", "Предложения по совершенствованию деятельности медицинских организаций, участвующих в реализации ТПОМС с целью повышения доступности и качества медицинской помощи" },
            { "Таблица 9", "Иная информация по вопросам реализации соглашения" },
        };

        public ReportDoffProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
            base(inClient, dgv, cmb, txtb, page,
                XmlFormTemplate.Doff.GetDescription(),
                Log,
                ReportGlobalConst.ReportDoff,
                reportsDictionary)
        {
            InitReport();
        }

        public override void InitReport()
        {
            Report = new ReportDoff { ReportDataList = new ReportDoffDto[ThemesList.Count], IdType = IdReportType };

            int i = 0;
            foreach (var theme in ThemesList.Select(x => x.Key))
            {
                Report.ReportDataList[i++] = new ReportDoffDto { Theme = theme };
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
                    reportType = ReportType.Doff
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response as ReportDoff;
        }

        public override void MapForAutoFill(AbstractReport report)
        {
            if (report == null)
            {
                return;
            }
            var inReport = report as ReportDoff;

            var index = Report.ReportDataList.ToList().FindIndex(x => x.Theme == Cmb.Text);
            var inTheme = inReport.ReportDataList.Single(x => x.Theme == Cmb.Text);
            Report.ReportDataList[index] = inTheme;
        }

        public override void FillDataGridView(string form)
        {
            if (form == null)
            {
                return;
            }
            if (_forms2346.Contains(form))
            {
                FillDgwForms2346(Dgv, form);
            }
            if (_forms31_41.Contains(form))
            {
                FillDgwForms31_41(Dgv, form);
            }



            if (Report.DataSource != DataSource.Handle)
            {
                //Dgv.DefaultCellStyle.BackColor = Color.LightGray;
            }
            else
            { Dgv.DefaultCellStyle.BackColor = Color.Azure; }
            //SetTotalColumn();
        }

        protected override void FillReport(string form)
        {
            if (form == null)
            {
                return;
            }
            if (_forms2346.Contains(form))
            {
                FillThemesForms2346(Dgv, form);
            }
            if (_forms31_41.Contains(form))
            {
                FillThemesForms31_41(Dgv, form);
            }

        }


        public override bool IsVisibleBtnDownloadExcel() => false;

        public override bool IsVisibleBtnHandle() => false;
        public override bool IsVisibleBtnSummary() => false;

        public override string ValidReport()
        {
            string message = "";
            return message;
        }

        public override void ToExcel(string filename, string filialName)
        {
            var excel = new ExcelDoffCreator(filename, ExcelForm.Doff, "", filialName, Report.Yymm);
            excel.CreateReport(Report, null);
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
                    reportType = ReportType.Doff
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportDoff;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
            Report.DataSource = response.DataSource;

        }

        public override void SaveReportDataSourceExcel()
        {
            var request = new SaveReportDataSourceExcelRequest
            {
                Body = new SaveReportDataSourceExcelRequestBody

                {
                    report = Report,
                    filialCode = CurrentUser.FilialCode,
                    idUser = CurrentUser.IdUser,
                    yymm = Report.Yymm,
                    reportType = ReportType.Doff
                }
            };
            var response = Client.SaveReportDataSourceExcel(request).Body.SaveReportDataSourceExcelResult as ReportDoff;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
            Report.DataSource = response.DataSource;

        }

        public override void SaveReportDataSourceHandle()
        {
            var request = new SaveReportDataSourceHandleRequest
            {
                Body = new SaveReportDataSourceHandleRequestBody

                {
                    report = Report,
                    filialCode = CurrentUser.FilialCode,
                    idUser = CurrentUser.IdUser,
                    yymm = Report.Yymm,
                    reportType = ReportType.Doff
                }
            };
            var response = Client.SaveReportDataSourceHandle(request).Body.SaveReportDataSourceHandleResult as ReportDoff;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
            Report.DataSource = response.DataSource;

        }

        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource)
        {
            var array = new ArrayOfString();
            array.AddRange(filialList);
            var request = new CollectSummaryReportRequest
            {
                Body = new CollectSummaryReportRequestBody
                {
                    filials = array,
                    status = status,
                    yymmStart = yymmStart,
                    yymmEnd = yymmEnd,
                    reportType = ReportType.Doff
                }
            };
            var response = Client.CollectSummaryReport(request);
            Report = response.Body.CollectSummaryReportResult as ReportDoff;
            Report.IdType = IdReportType;
            Report.Yymm = yymmEnd;
        }

        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            var formsList = ThemesList.Select(x => x.Key).OrderBy(x => x).ToList();
            var index = formsList.IndexOf(form);
            var currentHeaders = _headers[index];
            CreateDgvColumnsForTheme(Dgv, 400, _headersMap[form], currentHeaders);

            int countRows = ThemeTextData.Tables_fromxml.Single(x => x.TableName_fromxml == form).RowsCount_fromxml;
            foreach (var row in table)
            {
                var dgvRow = new DataGridViewRow();
                var cellName = new DataGridViewTextBoxCell
                {
                    Value = row.RowText_fromxml
                };
                var cellNum = new DataGridViewTextBoxCell
                {
                    Value = row.RowNum_fromxml
                };
                dgvRow.Cells.Add(cellName);
                dgvRow.Cells.Add(cellNum);

                var exclusionCells = row.ExclusionCells_fromxml?.Split(',');
                //for (int i = 2; i < countRows; i++)
                //{
                //    bool isNeedExcludeSum = exclusionCells?.Contains(i.ToString()) ?? false;
                //    var cell = new DataGridViewTextBoxCell
                //    {
                //        Value = row.Exclusion_fromxml || isNeedExcludeSum ? "x" : "0"
                //    };
                //    dgvRow.Cells.Add(cell);

                ////    if (isNeedExcludeSum)
                ////    {
                ////        cell.ReadOnly = true;
                ////        cell.Style.BackColor = Color.DarkGray;
                //    }
                //}
                int rowIndex = Dgv.Rows.Add(dgvRow);
                //if (row.Exclusion_fromxml)
                //{
                //    Dgv.Rows[rowIndex].ReadOnly = true;
                //    Dgv.Rows[rowIndex].DefaultCellStyle.BackColor = Color.LightCyan;
                //}

            }
        }

        private void CreateDgvColumnsForTheme(DataGridView dgvReport, int widthFirstColumn, string mainHeader,
            string[] columns)
        {
            if (columns.Length > 1)
            {
                CreateDgvCommonColumns(dgvReport, widthFirstColumn, mainHeader);
            }
            foreach (var column in columns)
            {
                var dgvColumn = new DataGridViewTextBoxColumn
                {
                    HeaderText = column,
                    Width = 100,
                    ReadOnly = false,
                    SortMode = DataGridViewColumnSortMode.NotSortable
                };
                dgvReport.Columns.Add(dgvColumn);
            }
        }

        private void CreateDgvCommonColumns(DataGridView dgvReport, int widthFirstColumn, string mainHeader)
        {
            dgvReport.AllowUserToAddRows = false;
            dgvReport.ColumnHeadersVisible = true;
            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = mainHeader,
                Width = widthFirstColumn,
                DataPropertyName = "Indicator",
                Name = "Indicator",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.Azure
                }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "№ строки",
                Width = 50,
                DataPropertyName = "NumRow",
                Name = "NumRow",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = Color.Azure
                }
            };
            dgvReport.Columns.Add(column);

        }

        private void FillThemesForms2346(DataGridView dgvReport, string form)
        {
            var reportDoffDto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportDoffDto != null)
            {
                reportDoffDto.Data = (from DataGridViewRow row in dgvReport.Rows
                                      let rowNum = row.Cells[1].Value.ToString().Trim()
                                      where !IsNotNeedFillRow(form, rowNum)
                                      select new ReportDoffDataDto
                                      {
                                          RowNum = rowNum,
                                          Column1 = GlobalUtils.TryParseInt(row.Cells[2].Value).ToString(),
                                          Column2 = GlobalUtils.TryParseInt(row.Cells[3].Value).ToString(),
                                          Column3 = GlobalUtils.TryParseInt(row.Cells[4].Value).ToString(),

                                      }).ToArray();
            }
        }

        private void FillDgwForms2346(DataGridView dgvReport, string form)
        {
            var reportDoffDto = Report.ReportDataList.FirstOrDefault(x => x.Theme == form);
            if (reportDoffDto == null)
            {
                return;
            }
            if (reportDoffDto.Data == null || reportDoffDto.Data.Length == 0)
            {
                return;
            }

            var rows = ThemeTextData.Tables_fromxml.Where(x => x.TableName_fromxml == form).SelectMany(x => x.Rows_fromxml).ToList();
            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                var rowNum = row.Cells[1].Value.ToString().Trim();
                //var exclusionsCells = rows.Single(x => x.RowNum_fromxml == rowNum).ExclusionCells_fromxml?.Split(',');
                //bool isExclusionsRow = rows.Single(x => x.RowNum_fromxml == rowNum).Exclusion_fromxml;

                var data = reportDoffDto.Data.SingleOrDefault(x => x.RowNum == rowNum);
                if (data != null)
                {
                    row.Cells[2].Value = data.Column1;
                    row.Cells[3].Value = data.Column2;
                    row.Cells[4].Value = data.Column3;

                }
            }
        }







        private void FillThemesForms31_41(DataGridView dgvReport, string form)
        {
            var reportDoffDto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportDoffDto != null)
            {
                reportDoffDto.Data = (from DataGridViewRow row in dgvReport.Rows
                                      let rowNum = row.Cells[1].Value.ToString().Trim()
                                      where !IsNotNeedFillRow(form, rowNum)
                                      select new ReportDoffDataDto
                                      {
                                          RowNum = rowNum,
                                          Column1 = GlobalUtils.TryParseInt(row.Cells[2].Value).ToString(),
                                          Column2 = GlobalUtils.TryParseInt(row.Cells[3].Value).ToString(),
                                          

                                      }).ToArray();
            }
        }

        private void FillDgwForms31_41(DataGridView dgvReport, string form)
        {
            var reportDoffDto = Report.ReportDataList.FirstOrDefault(x => x.Theme == form);
            if (reportDoffDto == null)
            {
                return;
            }
            if (reportDoffDto.Data == null || reportDoffDto.Data.Length == 0)
            {
                return;
            }

            var rows = ThemeTextData.Tables_fromxml.Where(x => x.TableName_fromxml == form).SelectMany(x => x.Rows_fromxml).ToList();
            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                var rowNum = row.Cells[1].Value.ToString().Trim();
                //var exclusionsCells = rows.Single(x => x.RowNum_fromxml == rowNum).ExclusionCells_fromxml?.Split(',');
                //bool isExclusionsRow = rows.Single(x => x.RowNum_fromxml == rowNum).Exclusion_fromxml;

                var data = reportDoffDto.Data.SingleOrDefault(x => x.RowNum == rowNum);
                if (data != null)
                {
                    row.Cells[2].Value = data.Column1;
                    row.Cells[3].Value = data.Column2;
                }
            }
        }




    }
}
