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
    class ReportZpz2025Processor : AbstractReportProcessor<ReportZpz2025>
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        private readonly string[] _forms9 = { "Таблица 9" };
        private readonly string[] _forms3 = { "Таблица 1" };
        private readonly string[] _forms1 = { "Таблица 10", "Таблица 4", "Таблица 8" };
        private readonly string[] _forms2 = { "Таблица 2", "Таблица 3" };
        private readonly string[] _forms67 = { "Таблица 6", "Таблица 7" };

        private readonly string[][] _headers = {
            new[]
            { "Устные", "Письменные","По поручениям" }, //1
            new[]
            {
                "разрешенные в досудебном порядке \r\n (гр.5)",
                "разрешенные в судебном порядке, в т.ч. по лицам, обратившимся за защитой прав ЗЛ: \r\n ЗЛ \r\n (гр.7)",
                "разрешенные в судебном порядке, в т.ч. по лицам, обратившимся за защитой прав ЗЛ: \r\n Представитель ЗЛ \r\n (гр.8)",
                "разрешенные в судебном порядке, в т.ч. по лицам, обратившимся за защитой прав ЗЛ: \r\n ТФОМС \r\n (гр.9)",
                "разрешенные в судебном порядке, в т.ч. по лицам, обратившимся за защитой прав ЗЛ: \r\n СМО \r\n (гр.10)",
                "разрешенные в судебном порядке, в т.ч. по лицам, обратившимся за защитой прав ЗЛ: \r\n Прокуратура \r\n (гр.11)"
            }, //2
            new[]
            {
                "разрешенные в досудебном порядке \r\n (гр.5)",
                "разрешенные в судебном порядке, в т.ч. по лицам, обратившимся за защитой прав ЗЛ: \r\n ЗЛ \r\n (гр.7)",
                "разрешенные в судебном порядке, в т.ч. по лицам, обратившимся за защитой прав ЗЛ: \r\n Представитель ЗЛ \r\n (гр.8)",
                "разрешенные в судебном порядке, в т.ч. по лицам, обратившимся за защитой прав ЗЛ: \r\n ТФОМС \r\n (гр.9)",
                "разрешенные в судебном порядке, в т.ч. по лицам, обратившимся за защитой прав ЗЛ: \r\n СМО \r\n (гр.10)",
                "разрешенные в судебном порядке, в т.ч. по лицам, обратившимся за защитой прав ЗЛ: \r\n Прокуратура \r\n (гр.11)"
            }, //3
            new[]
            { "Всего" }, //4
            new[]
            {
                "внеплановая МЭЭ вне медицинской организации \r\n (гр.4)",
                "внеплановая МЭЭ амбулаторно \r\n (гр.5)",
                "внеплановая МЭЭ в дневном стационаре \r\n (гр.6)",
                "внеплановая МЭЭ в том числе ВМП \r\n (гр.7)",
                "внеплановая МЭЭ стационарно \r\n (гр.8)",
                "внеплановая МЭЭ в том числе ВМП \r\n (гр.9)",
                "плановая МЭЭ вне медицинской организации \r\n (гр.11)",
                "плановая МЭЭ амбулаторно \r\n (гр.12)",
                "плановая МЭЭ в дневном стационаре \r\n (гр.13)",
                "плановая МЭЭ в том числе ВМП \r\n (гр.14)",
                "плановая МЭЭ стационарно \r\n (гр.15)",
                "плановая МЭЭ в том числе ВМП \r\n (гр.16)"
            }, //6
            new[]
            {
                "целевая ЭКМП вне медицинской организации \r\n (гр.4)",
                "целевая ЭКМП амбулаторно \r\n (гр.5)",
                "целевая ЭКМП в дневном стационаре \r\n (гр.6)",
                "целевая ЭКМП в том числе ВМП \r\n (гр.7)",
                "целевая ЭКМП стационарно \r\n (гр.8)",
                "целевая ЭКМП в том числе ВМП \r\n (гр.9)",
                "плановая ЭКМП вне медицинской организации \r\n (гр.11)",
                "плановая ЭКМП амбулаторно \r\n (гр.12)",
                "плановая ЭКМП в дневном стационаре \r\n (гр.13)",
                "плановая ЭКМП в том числе ВМП \r\n (гр.14)",
                "плановая ЭКМП стационарно \r\n (гр.15)",
                "плановая ЭКМП в том числе ВМП \r\n (гр.16)"
            }, //7
            new[]
            { "Всего" }, //8
            new[]
            { "штатные работники\r\n (гр.7)", "привлекаемые \r\nпо гражданско-правовому договору \r\n (гр.9)" }, //9

        };

        private readonly Dictionary<string, string> _headersMap = new Dictionary<string, string>
        {
            { "Таблица 1", "Виды обращений" },
            { "Таблица 2", "Количество спорных случаев (сумма возмещения ущерба, причиненного застрахованным лицам)" },
            { "Таблица 3", "Виды обращений" },
            { "Таблица 4", "Количество исков в порядке регресса (сумма средств, полученных по регрессным искам)" },
            { "Таблица 6", "Количество проведенных медико-экономических экспертиз медицинской помощи (далее - МЭЭ) (выявленных нарушений)" },
            { "Таблица 7", "Количество проведенных экспертиз качества медицинской помощи (далее - ЭКМП) (выявленных нарушений)" },
            { "Таблица 8", "Финансовые результаты" },
            { "Таблица 9", "Специалисты, участвующие в защите прав застрахованных лиц" },
        };

        public ReportZpz2025Processor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
            base(inClient, dgv, cmb, txtb, page,
                XmlFormTemplate.Zpz2025.GetDescription(),
                Log,
                ReportGlobalConst.ReportZpz2025,
                reportsDictionary)
        {
            InitReport();
        }

        public override void InitReport()
        {
            Report = new ReportZpz2025 { ReportDataList = new ReportZpz2025Dto[ThemesList.Count], IdType = IdReportType };

            int i = 0;
            foreach (var theme in ThemesList.Select(x => x.Key))
            {
                Report.ReportDataList[i++] = new ReportZpz2025Dto { Theme = theme };
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
                    reportType = ReportType.Zpz2025
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response as ReportZpz2025;
        }

        public override void MapForAutoFill(AbstractReport report)
        {
            if (report == null)
            {
                return;
            }
            var inReport = report as ReportZpz2025;

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
            if (_forms3.Contains(form))
            {
                FillDgwForms3(Dgv, form);
            }
            else if (_forms9.Contains(form))
            {
                FillDgwForms9(Dgv, form);
            }
            else if (_forms1.Contains(form))
            {
                FillDgwForms1(Dgv, form);
            }
            else if (_forms67.Contains(form))
            {
                FillDgwForms67(Dgv, form);
            }
            else if (_forms2.Contains(form))
            {
                FillDgwForms2(Dgv, form);
            }
            else
            {
                FillDgwForms5(Dgv, form);
            }

            if (Report.DataSource != DataSource.Handle)
            {
                Dgv.DefaultCellStyle.BackColor = Color.LightGray;
            }
            else
            {
                Dgv.DefaultCellStyle.BackColor = Color.Azure;
            }
            SetTotalColumn();
        }

        protected override void FillReport(string form)
        {
            if (form == null)
            {
                return;
            }
            if (_forms3.Contains(form))
            {
                FillThemesForms3(Dgv, form);
            }
            else if (_forms9.Contains(form))
            {
                FillThemesForms9(Dgv, form);
            }
            else if (_forms2.Contains(form))
            {
                FillThemesForms2(Dgv, form);
            }
            else if (_forms1.Contains(form))
            {
                FillThemesForms1(Dgv, form);
            }
            else if (_forms67.Contains(form))
            {
                FillThemesForms67(Dgv, form);
            }
            else
            {
                FillThemesForms5(Dgv, form);
            }
        }

        public override bool IsVisibleBtnDownloadExcel() => true;
        public override bool IsVisibleBtnHandle() => true;
        public override bool IsVisibleBtnSummary() => false;

        public override string ValidReport()
        {
            string message = "";
            string[] validForms = { "Таблица 6", "Таблица 7" };
            foreach (var data in Report.ReportDataList.Where(x => validForms.Contains(x.Theme)))
            {
                if (data.Data == null)
                {
                    continue;
                }

                string localMessage = "";
                string lastSumRow = data.Theme == "Таблица 6" ? "2.5, 2.6" : "2.5";

                decimal gr4 = data.Data.Where(x => x.Code == "2").Sum(x => x.CountOutOfSmo);
                decimal gr4Another = data.Data.Where(x => x.Code.StartsWith("2") && x.Code.Length == 3).Sum(x => x.CountOutOfSmo);
                decimal gr5 = data.Data.Where(x => x.Code == "2").Sum(x => x.CountAmbulatory);
                decimal gr5Another = data.Data.Where(x => x.Code.StartsWith("2") && x.Code.Length == 3).Sum(x => x.CountAmbulatory);
                decimal gr6 = data.Data.Where(x => x.Code == "2").Sum(x => x.CountDs);
                decimal gr6Another = data.Data.Where(x => x.Code.StartsWith("2") && x.Code.Length == 3).Sum(x => x.CountDs);
                decimal gr7 = data.Data.Where(x => x.Code == "2").Sum(x => x.CountDsVmp);
                decimal gr7Another = data.Data.Where(x => x.Code.StartsWith("2") && x.Code.Length == 3).Sum(x => x.CountDsVmp);
                decimal gr8 = data.Data.Where(x => x.Code == "2").Sum(x => x.CountStac);
                decimal gr8Another = data.Data.Where(x => x.Code.StartsWith("2") && x.Code.Length == 3).Sum(x => x.CountStac);
                decimal gr9 = data.Data.Where(x => x.Code == "2").Sum(x => x.CountStacVmp);
                decimal gr9Another = data.Data.Where(x => x.Code.StartsWith("2") && x.Code.Length == 3).Sum(x => x.CountStacVmp);
                decimal gr11 = data.Data.Where(x => x.Code == "2").Sum(x => x.CountOutOfSmoAnother);
                decimal gr11Another = data.Data.Where(x => x.Code.StartsWith("2") && x.Code.Length == 3).Sum(x => x.CountOutOfSmoAnother);

                if (gr4 < gr4Another)
                {
                    localMessage += $"В строке 2 сумма гр.4 должна быть больше или равна сумме строк {lastSumRow} гр.4\n";
                }
                if (gr5 < gr5Another)
                {
                    localMessage += $"В строке 2 сумма гр.5 должна быть больше или равна сумме строк {lastSumRow} гр.5\n";
                }
                if (gr6 < gr6Another)
                {
                    localMessage += $"В строке 2 сумма гр.6 должна быть больше или равна сумме строк {lastSumRow} гр.6\n";
                }
                if (gr7 < gr7Another)
                {
                    localMessage += $"В строке 2 сумма гр.7 должна быть больше или равна сумме строк {lastSumRow} гр.7\n";
                }
                if (gr8 < gr8Another)
                {
                    localMessage += $"В строке 2 сумма гр.8 должна быть больше или равна сумме строк {lastSumRow} гр.8\n";
                }
                if (gr9 < gr9Another)
                {
                    localMessage += $"В строке 2 сумма гр.9 должна быть больше или равна сумме строк {lastSumRow} гр.9\n";
                }
                if (gr11 < gr11Another)
                {
                    localMessage += $"В строке 2 сумма гр.11 должна быть больше или равна сумме строк {lastSumRow} гр.11\n";
                }

                if (!string.IsNullOrEmpty(localMessage))
                {
                    message += $"Тема {data.Theme}\n" + localMessage;
                }
            }

            return message;
        }

        public override void ToExcel(string filename, string filialName)
        {
            var mm = YymmUtils.GetMonth(Report.Yymm.Substring(2, 2)) + " 20" + Report.Yymm.Substring(0, 2);
            var excel = new ExcelZpz2025Creator(filename, ExcelForm.Zpz2025, mm, filialName);
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
                    reportType = ReportType.Zpz2025
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportZpz2025;
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
                    reportType = ReportType.Zpz2025
                }
            };
            var response = Client.SaveReportDataSourceExcel(request).Body.SaveReportDataSourceExcelResult as ReportZpz2025;
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
                    reportType = ReportType.Zpz2025
                }
            };
            var response = Client.SaveReportDataSourceHandle(request).Body.SaveReportDataSourceHandleResult as ReportZpz2025;
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
                    reportType = ReportType.Zpz2025
                }
            };
            var response = Client.CollectSummaryReport(request);
            Report = response.Body.CollectSummaryReportResult as ReportZpz2025;
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
                for (int i = 2; i < countRows; i++)
                {
                    bool isNeedExcludeSum = exclusionCells?.Contains(i.ToString()) ?? false;
                    var cell = new DataGridViewTextBoxCell
                    {
                        Value = row.Exclusion_fromxml || isNeedExcludeSum ? "x" : "0"
                    };
                    dgvRow.Cells.Add(cell);

                    if (isNeedExcludeSum)
                    {
                        cell.ReadOnly = true;
                        cell.Style.BackColor = Color.DarkGray;
                    }
                }
                int rowIndex = Dgv.Rows.Add(dgvRow);
                if (row.Exclusion_fromxml)
                {
                    Dgv.Rows[rowIndex].ReadOnly = true;
                    Dgv.Rows[rowIndex].DefaultCellStyle.BackColor = Color.LightCyan;
                }

            }
        }

        private void CreateDgvColumnsForTheme(DataGridView dgvReport, int widthFirstColumn, string mainHeader,
            string[] columns)
        {
            CreateDgvCommonColumns(dgvReport, widthFirstColumn, mainHeader);
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

        private void FillThemesForms2(DataGridView dgvReport, string form)
        {
            var reportZpz2025Dto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportZpz2025Dto != null)
            {
                reportZpz2025Dto.Data = (from DataGridViewRow row in dgvReport.Rows
                                         let rowNum = row.Cells[1].Value.ToString().Trim()
                                         where !IsNotNeedFillRow(form, rowNum)
                                         select new ReportZpz2025DataDto
                                         {
                                             Code = rowNum,
                                             CountSmo = GlobalUtils.TryParseDecimal(row.Cells[2].Value),
                                             CountInsured = GlobalUtils.TryParseDecimal(row.Cells[3].Value),
                                             CountInsuredRepresentative = GlobalUtils.TryParseDecimal(row.Cells[4].Value),
                                             CountTfoms = GlobalUtils.TryParseDecimal(row.Cells[5].Value),
                                             CountSmoAnother = GlobalUtils.TryParseDecimal(row.Cells[6].Value),
                                             CountProsecutor = GlobalUtils.TryParseDecimal(row.Cells[7].Value)
                                         }).ToArray();
            }
        }

        private void FillThemesForms3(DataGridView dgvReport, string form)
        {
            var reportZpz2025Dto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportZpz2025Dto == null)
            {
                return;
            }

            reportZpz2025Dto.Data = (from DataGridViewRow row in dgvReport.Rows
                                     let rowNum = row.Cells[1].Value.ToString().Trim()
                                     where !IsNotNeedFillRow(form, rowNum)
                                     select new ReportZpz2025DataDto
                                     {
                                         Code = rowNum,
                                         CountSmo = GlobalUtils.TryParseDecimal(row.Cells[2].Value),
                                         CountSmoAnother = GlobalUtils.TryParseDecimal(row.Cells[3].Value),
                                         CountAssignment = GlobalUtils.TryParseDecimal(row.Cells[4].Value)
                                     }).ToArray();
        }

        private void FillThemesForms9(DataGridView dgvReport, string form)
        {
            var reportZpz2025Dto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportZpz2025Dto == null)
            {
                return;
            }

            reportZpz2025Dto.Data = (from DataGridViewRow row in dgvReport.Rows
                                     let rowNum = row.Cells[1].Value.ToString().Trim()
                                     where !IsNotNeedFillRow(form, rowNum)
                                     select new ReportZpz2025DataDto
                                     {
                                         Code = rowNum,
                                         CountSmo = GlobalUtils.TryParseDecimal(row.Cells[2].Value),
                                         CountSmoAnother = GlobalUtils.TryParseDecimal(row.Cells[3].Value),
                                         CountAssignment = GlobalUtils.TryParseDecimal(row.Cells[4].Value)
                                     }).ToArray();
        }

        private void FillThemesForms1(DataGridView dgvReport, string form)
        {
            var reportZpz2025Dto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportZpz2025Dto == null)
            {
                return;
            }

            reportZpz2025Dto.Data = (from DataGridViewRow row in dgvReport.Rows
                                     let rowNum = row.Cells[1].Value.ToString().Trim()
                                     where !IsNotNeedFillRow(form, rowNum)
                                     select new ReportZpz2025DataDto
                                     {
                                         Code = rowNum,
                                         CountSmo = GlobalUtils.TryParseDecimal(row.Cells[2].Value),
                                         CountSmoAnother = GlobalUtils.TryParseDecimal(row.Cells[3].Value)
                                     }).ToArray();
        }



        private void FillThemesForms67(DataGridView dgvReport, string form)
        {

                     
            var reportZpz2025Dto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportZpz2025Dto != null)
            {
                reportZpz2025Dto.Data = (from DataGridViewRow row in dgvReport.Rows
                                    let rowNum = row.Cells[1].Value.ToString().Trim()
                                    where !IsNotNeedFillRow(form, rowNum)
                                    select new ReportZpz2025DataDto
                                    {
                                        Code = rowNum,
                                        CountOutOfSmo = GlobalUtils.TryParseDecimal(row.Cells[2].Value),
                                        CountAmbulatory = GlobalUtils.TryParseDecimal(row.Cells[3].Value),
                                        CountDs = GlobalUtils.TryParseDecimal(row.Cells[4].Value),
                                        CountDsVmp = GlobalUtils.TryParseDecimal(row.Cells[5].Value),
                                        CountStac = GlobalUtils.TryParseDecimal(row.Cells[6].Value),
                                        CountStacVmp = GlobalUtils.TryParseDecimal(row.Cells[7].Value),
                                        CountOutOfSmoAnother = GlobalUtils.TryParseDecimal(row.Cells[8].Value),
                                        CountAmbulatoryAnother = GlobalUtils.TryParseDecimal(row.Cells[9].Value),
                                        CountDsAnother = GlobalUtils.TryParseDecimal(row.Cells[10].Value),
                                        CountDsVmpAnother = GlobalUtils.TryParseDecimal(row.Cells[11].Value),
                                        CountStacAnother = GlobalUtils.TryParseDecimal(row.Cells[12].Value),
                                        CountStacVmpAnother = GlobalUtils.TryParseDecimal(row.Cells[13].Value)
                                    }).ToArray();

                 
            }
        }

        private void FillThemesForms5(DataGridView dgvReport, string form)
        {
            var reportZpz2025Dto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportZpz2025Dto != null)
            {
                reportZpz2025Dto.Data = (from DataGridViewRow row in dgvReport.Rows
                                    let rowNum = row.Cells[1].Value.ToString().Trim()
                                    where !IsNotNeedFillRow(form, rowNum)
                                    select new ReportZpz2025DataDto
                                    {
                                        Code = rowNum,
                                        CountOutOfSmo = GlobalUtils.TryParseDecimal(row.Cells[2].Value),
                                        CountAmbulatory = GlobalUtils.TryParseDecimal(row.Cells[3].Value),
                                        CountDs = GlobalUtils.TryParseDecimal(row.Cells[4].Value),
                                        CountDsVmp = GlobalUtils.TryParseDecimal(row.Cells[5].Value),
                                        CountStac = GlobalUtils.TryParseDecimal(row.Cells[6].Value),
                                        CountStacVmp = GlobalUtils.TryParseDecimal(row.Cells[7].Value)
                                    }).ToArray();
            }
        }

        private void FillDgwForms3(DataGridView dgvReport, string form)
        {
            var reportZpz2025Dto = Report.ReportDataList?.Single(x => x.Theme == form);
            if (reportZpz2025Dto?.Data == null || reportZpz2025Dto.Data.Length == 0)
            {
                return;
            }

            var rows = ThemeTextData.Tables_fromxml.Where(x => x.TableName_fromxml == form).SelectMany(x => x.Rows_fromxml).ToList();
            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                var rowNum = row.Cells[1].Value.ToString().Trim();
                bool exclusionsRow = rows.Single(x => x.RowNum_fromxml == rowNum).Exclusion_fromxml;

                var data = reportZpz2025Dto.Data.SingleOrDefault(x => x.Code == rowNum);
                if (data != null)
                {
                    row.Cells[2].Value = exclusionsRow ? "x" : data.CountSmo.ToString().Replace(",00", "");
                    row.Cells[3].Value = exclusionsRow ? "x" : data.CountSmoAnother.ToString().Replace(",00", "");
                    row.Cells[4].Value = exclusionsRow ? "x" : data.CountAssignment.ToString().Replace(",00", "");
                }
            }
        }

        //private void FillDgwForms1(DataGridView dgvReport, string form)
        //{
        //    var reportZpz2025Dto = Report.ReportDataList.FirstOrDefault(x => x.Theme == form);
        //    if (reportZpz2025Dto?.Data == null || reportZpz2025Dto?.Data?.Length == 0)
        //    {
        //        return;
        //    }

        //    var rows = ThemeTextData.tables.Where(x => x.Name == form).SelectMany(x => x.Rows).ToList();
        //    foreach (DataGridViewRow row in dgvReport.Rows)
        //    {
        //        var rowNum = row.Cells[1].Value.ToString().Trim();
        //        bool exclusionsRow = rows.Single(x => x.Num == rowNum).Exclusion;
        //        var data = reportZpz2025Dto.Data.SingleOrDefault(x => x.Code == rowNum);
        //        if (data != null)
        //        {
        //            row.Cells[2].Value = exclusionsRow ? "x" : data.CountSmo.ToString().Replace(",00", "");
        //        }
        //    }

        //}

        private void FillDgwForms9(DataGridView dgvReport, string form)
        {
            var reportZpz2025Dto = Report.ReportDataList?.Single(x => x.Theme == form);
            if (reportZpz2025Dto?.Data == null || reportZpz2025Dto.Data.Length == 0)
            {
                return;
            }

            var rows = ThemeTextData.Tables_fromxml.Where(x => x.TableName_fromxml == form).SelectMany(x => x.Rows_fromxml).ToList();
            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                var rowNum = row.Cells[1].Value.ToString().Trim();
                bool isExclusionsRow = rows.Single(x => x.RowNum_fromxml == rowNum).Exclusion_fromxml;

                var data = reportZpz2025Dto.Data.SingleOrDefault(x => x.Code == rowNum);
                if (data != null)
                {
                    row.Cells[2].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.CountSmo);
                    row.Cells[3].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.CountSmoAnother);
                    row.Cells[4].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.CountAssignment);

                }
            }
        }

        private void FillDgwForms1(DataGridView dgvReport, string form)
        {
            var reportZpz2025Dto = Report.ReportDataList?.Single(x => x.Theme == form);
            if (reportZpz2025Dto?.Data == null || reportZpz2025Dto.Data.Length == 0)
            {
                return;
            }

            var rows = ThemeTextData.Tables_fromxml.Where(x => x.TableName_fromxml == form).SelectMany(x => x.Rows_fromxml).ToList();
            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                var rowNum = row.Cells[1].Value.ToString().Trim();
                bool isExclusionsRow = rows.Single(x => x.RowNum_fromxml == rowNum).Exclusion_fromxml;

                var data = reportZpz2025Dto.Data.SingleOrDefault(x => x.Code == rowNum);
                if (data != null)
                {
                    row.Cells[2].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.CountSmo);
                    row.Cells[3].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.CountSmoAnother);
                }
            }
        }

        private void FillDgwForms2(DataGridView dgvReport, string form)
        {
            var reportZpz2025Dto = Report.ReportDataList.FirstOrDefault(x => x.Theme == form);
            if (reportZpz2025Dto == null)
            {
                return;
            }
            if (reportZpz2025Dto.Data == null || reportZpz2025Dto.Data.Length == 0)
            {
                return;
            }

            var rows = ThemeTextData.Tables_fromxml.Where(x => x.TableName_fromxml == form).SelectMany(x => x.Rows_fromxml).ToList();
            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                var rowNum = row.Cells[1].Value.ToString().Trim();
                var exclusionsCells = rows.Single(x => x.RowNum_fromxml == rowNum).ExclusionCells_fromxml?.Split(',');
                bool isExclusionsRow = rows.Single(x => x.RowNum_fromxml == rowNum).Exclusion_fromxml;

                var data = reportZpz2025Dto.Data.SingleOrDefault(x => x.Code == rowNum);
                if (data != null)
                {
                    row.Cells[2].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 2, data.CountSmo);
                    row.Cells[3].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 3, data.CountInsured);
                    row.Cells[4].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 4, data.CountInsuredRepresentative);
                    row.Cells[5].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 5, data.CountTfoms);
                    row.Cells[6].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 6, data.CountSmoAnother);
                    row.Cells[7].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 7, data.CountProsecutor);
                }
            }
        }




        private void FillDgwForms67(DataGridView dgvReport, string form)
        {
            var reportZpz2025Dto = Report.ReportDataList.Single(x => x.Theme == form);
            if (reportZpz2025Dto.Data == null || reportZpz2025Dto.Data.Length == 0)
            {
                return;
            }

            var rows = ThemeTextData.Tables_fromxml.Where(x => x.TableName_fromxml == form).SelectMany(x => x.Rows_fromxml).ToList();
            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                var rowNum = row.Cells[1].Value.ToString().Trim();
                var exclusionsCells = rows.Single(x => x.RowNum_fromxml == rowNum).ExclusionCells_fromxml?.Split(',');
                bool isExclusionsRow = rows.Single(x => x.RowNum_fromxml == rowNum).Exclusion_fromxml;

                var data = reportZpz2025Dto.Data.SingleOrDefault(x => x.Code == rowNum);
                if (data == null)
                {
                    continue;
                }

                row.Cells[2].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 2, data.CountOutOfSmo);
                row.Cells[3].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 3, data.CountAmbulatory);
                row.Cells[4].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 4, data.CountDs);             
                row.Cells[5].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 5, data.CountDsVmp);
                row.Cells[6].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 6, data.CountStac);
                row.Cells[7].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 7, data.CountStacVmp);
                row.Cells[8].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 8, data.CountOutOfSmoAnother);
                row.Cells[9].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 9, data.CountAmbulatoryAnother);
                row.Cells[10].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 10, data.CountDsAnother);
                row.Cells[11].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 11, data.CountDsVmpAnother);
                row.Cells[12].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 12, data.CountStacAnother);
                row.Cells[13].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, exclusionsCells, 13, data.CountStacVmpAnother);
            }

        
        }

        private void FillDgwForms5(DataGridView dgvReport, string form)
        {
            var reportZpz2025Dto = Report.ReportDataList.Single(x => x.Theme == form);
            if (reportZpz2025Dto.Data == null || reportZpz2025Dto.Data.Length == 0)
            {
                return;
            }

            var rows = ThemeTextData.Tables_fromxml.Where(x => x.TableName_fromxml == form).SelectMany(x => x.Rows_fromxml).ToList();
            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                var rowNum = row.Cells[1].Value.ToString().Trim();
                var data = reportZpz2025Dto.Data.SingleOrDefault(x => x.Code == rowNum);
                bool isExclusionsRow = rows.Single(x => x.RowNum_fromxml == rowNum).Exclusion_fromxml;
                if (data == null)
                {
                    continue;
                }

                row.Cells[2].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.CountOutOfSmo);
                row.Cells[3].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.CountAmbulatory);
                row.Cells[4].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.CountDs);
                row.Cells[5].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.CountDsVmp);
                row.Cells[6].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.CountStac);
                row.Cells[7].Value = ZpzDgvUtils.GetRowText(isExclusionsRow, null, 0, data.CountStacVmp);
            }

        }

    }

}
