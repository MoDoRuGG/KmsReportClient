using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.XML;
using KmsReportClient.Support;
using NLog;

namespace KmsReportClient.Report.Basic
{
    internal class ReportDoffProcessor : AbstractReportProcessor<ReportDoff>
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        string[] _notSaveCells = new string[] { "1.", "1.а)", "1.б)", "1.в)", "1.1.", "1.2." };

        private readonly string[] _forms2 = { "Таблица 2" };
        private readonly string[] _forms346 = { "Таблица 3", "Таблица 4", "Таблица 6" };
        private readonly string[] _forms159 = { "Таблица 1", "Таблица 5", "Таблица 9" };
        private readonly string[] _forms3141 = { "Таблица 3.1", "Таблица 4.1" };
        private readonly string[] _forms78 = { "Таблица 7", "Таблица 8" };

        Dictionary<string, DataGridViewRow> _rows;


        public ReportDoffProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
            base(inClient, dgv, cmb, txtb, page,
                XmlFormTemplate.Doff.GetDescription(),
                Log,
                ReportGlobalConst.ReportDoff,
                reportsDictionary)
        {
            InitReport();
        }


        private void SetStyle()
        {
            foreach (DataGridViewRow row in Dgv.Rows)
            {
                string rowNum = row.Cells[1].Value.ToString();
                if (_notSaveCells.Contains(rowNum))
                {
                    row.DefaultCellStyle.BackColor = Color.LightGray;

                    row.ReadOnly = true;
                    row.DefaultCellStyle.Font = new Font(Dgv.DefaultCellStyle.Font, FontStyle.Bold);
                }
                row.Cells[3].Style.BackColor = Color.LightGray;
                row.Cells[4].Style.BackColor = Color.LightGray;
            }
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
            return response == null ? null : response as ReportDoff;
        }

        public override void FillDataGridView(string form)
        {
            Dgv.RowHeadersWidth = 30;
            if (form == null)
            {
                return;
            }


            else if (_forms159.Contains(form))
            {
                FillDgvForms159(Dgv, form);
            }
            else if (_forms2.Contains(form))
            {
                FillDgvForms2(Dgv, form);
            }            
            else if (_forms346.Contains(form))
            {
                FillDgvForms346(Dgv, form);
            }

        }

        protected override void FillReport(string form)
        {
            if (form == null)
            {
                return;
            }
            if (_forms159.Contains(form))
            {
                FillThemesForms159(Dgv, form);
            }
            else if (_forms2.Contains(form))
            {
                FillThemesForms2(Dgv, form);
            }
            else if (_forms346.Contains(form))
            {
                FillThemesForms346(Dgv, form);
            }

        }


        public override void ToExcel(string filename, string filialName)
        {
            //var excel = new ExcelDoffCreator(filename, ExcelForm.Doff, Report.Yymm, filialName);
            //var yearReport = FillYearReport();
            //excel.CreateReport(Report, yearReport);
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
        }

        public override void SaveReportDataSourceExcel()
        { }
        public override void SaveReportDataSourceHandle()
        { }

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
            Report = Client.CollectSummaryReport(request).Body.CollectSummaryReportResult as ReportDoff;
            Report.IdType = IdReportType;
            Report.Yymm = yymmEnd;
        }

        public override bool IsVisibleBtnDownloadExcel() => false;

        public override bool IsVisibleBtnHandle()
        {
            return false;
        }

        public override bool IsVisibleBtnSummary() => false;

        public override string ValidReport()
        {
            var message = "";
            //foreach (var data in Report.ReportDataList)
            //{
            //    if (data.Theme == "Таблица 3")
            //    {
            //        var i = 1;
            //        if (data.Table3 == null || data.Table3.Length == 0)
            //        {
            //            message = @"Перед выгрузкой в Excel необходимо заполнить таблицу 3";
            //            continue;
            //        }

            //        foreach (var table3 in data.Table3)
            //        {
            //            if (string.IsNullOrEmpty(table3.Mo))
            //            {
            //                message = $"Строка {i}. гр.1 не может быть пустым; ";
            //            }
            //            if (table3.CountUnit < 1)
            //            {
            //                message += $"Строка {i}. гр.2 должна быть больше 0" + Environment.NewLine;
            //            }
            //            if (table3.CountUnitChild > table3.CountUnit)
            //            {
            //                message += $"Строка {i}. гр.3 должно быть меньше или равна гр.2" + Environment.NewLine;
            //            }
            //            if (table3.CountUnitWithSpChild > table3.CountUnitWithSp)
            //            {
            //                message += $"Строка {i}. гр.5 должно быть меньше или равна гр.4" + Environment.NewLine;
            //            }
            //            if (table3.CountChannelSpChild > table3.CountChannelSp)
            //            {
            //                message += $"Строка {i}. гр.7 должно быть меньше или равна гр.6" + Environment.NewLine;
            //            }
            //            if (table3.CountChannelPhoneChild > table3.CountChannelPhone)
            //            {
            //                message += $"Строка {i}. гр.9 должно быть меньше или равна гр.8" + Environment.NewLine;
            //            }
            //            if (table3.CountChannelTerminalChild > table3.CountChannelTerminal)
            //            {
            //                message += $"Строка {i}. гр.11 должно быть меньше или равна гр.10" + Environment.NewLine;
            //            }
            //            if (table3.CountChannelAnotherChild > table3.CountChannelAnother)
            //            {
            //                message += $"Строка {i}. гр.13 должно быть меньше или равна гр.12" + Environment.NewLine;
            //            }
            //            if (table3.CountUnitWithSp > table3.CountUnit)
            //            {
            //                message += $"Строка {i}. гр.4 должно быть меньше или равна гр.2" + Environment.NewLine;
            //            }
            //            if (table3.CountUnitWithSpChild > table3.CountUnitChild)
            //            {
            //                message += $"Строка {i}. гр.5 должно быть меньше или равна гр.3" + Environment.NewLine;
            //            }

            //            var sumChild =
            //                table3.CountChannelSpChild +
            //                table3.CountChannelPhoneChild +
            //                table3.CountChannelTerminalChild +
            //                table3.CountChannelAnotherChild;
            //            if (table3.CountUnitWithSpChild > sumChild)
            //            {
            //                message += $"Строка {i}. гр.5 не должна быть больше суммы гр.7, 9, 11, 13" + Environment.NewLine;
            //            }

            //            if (table3.CountUnitWithSpChild < 0)
            //            {
            //                message += $"Строка {i}. гр.5 не должна быть меньше 0" + Environment.NewLine;
            //            }

            //            var sum = table3.CountChannelSp +
            //                      table3.CountChannelPhone +
            //                      table3.CountChannelTerminal +
            //                      table3.CountChannelAnother;
            //            if (table3.CountUnitWithSp > sum)
            //            {
            //                message += $"Строка {i}. гр.4 не должна быть больше суммы гр. 6, 8, 10, 12" + Environment.NewLine;
            //            }
            //            if (table3.CountUnitWithSp < 0)
            //            {
            //                message += $"Строка {i}. гр.4 не должна быть меньше 0" + Environment.NewLine;
            //            }

            //            i++;
            //        }

            //        if (message.Length > 0)
            //        {
            //            message = "Форма 262. Таблица 3. " + Environment.NewLine + message;

            //            Dgv.RowHeadersWidth = 60;
            //            foreach (DataGridViewRow row in Dgv.Rows)
            //            {
            //                row.HeaderCell.Value = Convert.ToString(row.Index + 1);
            //            }
            //        }
            //    }
            //}

            return message;
        }

        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            if (form == "Таблица 1" || form == "Таблица 5" || form == "Таблица 9")
            {
                CreateDgvColumnsForTable159(Dgv);
                foreach (var row in table)
                {
                    Dgv.Rows.Add("");
                }
            }
            else if (form == "Таблица 3.1")
            {
                CreateDgvColumnsForTable3_1(Dgv);
                foreach (var row in table)
                {
                    Dgv.Rows.Add("");
                }
            }
            else if (form == "Таблица 4.1")
            {
                CreateDgvColumnsForTable4_1(Dgv);
                foreach (var row in table)
                {
                    Dgv.Rows.Add("");
                }
            }
            else if (form == "Таблица 2")
            {
                CreateDgvColumnsForTable2(Dgv);
                foreach (var row in table)
                {
                    Dgv.Rows.Add(row.RowText_fromxml, row.RowNum_fromxml, "", "", "");

                }
                SetStyle();
                _rows = new Dictionary<string, DataGridViewRow>();
                foreach (DataGridViewRow row in Dgv.Rows)
                {
                    _rows.Add(row.Cells[1].Value.ToString(), row);
                }
            }
            else if (form == "Таблица 3")
            {
                CreateDgvColumnsForTable3(Dgv);
                foreach (var row in table)
                {
                    Dgv.Rows.Add(row.RowText_fromxml, row.RowNum_fromxml, "", "", "");
                }
                SetStyle();
                _rows = new Dictionary<string, DataGridViewRow>();
                foreach (DataGridViewRow row in Dgv.Rows)
                {
                    _rows.Add(row.Cells[1].Value.ToString(), row);
                }
            }
            else if (form == "Таблица 4")
            {
                CreateDgvColumnsForTable4(Dgv);
                foreach (var row in table)
                {
                    Dgv.Rows.Add(row.RowText_fromxml, row.RowNum_fromxml, "", "", "");
                }
                SetStyle();
                _rows = new Dictionary<string, DataGridViewRow>();
                foreach (DataGridViewRow row in Dgv.Rows)
                {
                    _rows.Add(row.Cells[1].Value.ToString(), row);
                }
            }
            else if (form == "Таблица 6")
            {
                CreateDgvColumnsForTable6(Dgv);
                foreach (var row in table)
                {
                    Dgv.Rows.Add(row.RowText_fromxml, row.RowNum_fromxml, "", "", "");
                }
                SetStyle();
                _rows = new Dictionary<string, DataGridViewRow>();
                foreach (DataGridViewRow row in Dgv.Rows)
                {
                    _rows.Add(row.Cells[1].Value.ToString(), row);
                }
            }
            else if (form == "Таблица 7")
            {
                Dgv.AllowUserToAddRows = true;
                CreateDgvColumnsForTable7(Dgv);
            }
            else if (form == "Таблица 8")
            {
                Dgv.AllowUserToAddRows = true;
                CreateDgvColumnsForTable8(Dgv);

            }


        }

        private void CreateDgvColumnsForTable7(DataGridView dgvReport)
        {
            dgvReport.AllowUserToAddRows = true;
            dgvReport.ColumnHeadersVisible = true;
            var headerText1 =
                "№ п/п";
            var headerText2 =
                "Дата проведения (дд.мм.гггг)";
            var headerText3 =
                "Вид мероприятия, наименование, инициатор (руководитель мероприятия)";
            var headerText4 =
                "Краткое содержание, относящееся к тематике соглашения (участники, темы обсуждения, решения)";

            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = headerText1,
                Width = 50,
                DataPropertyName = "RowNum",
                Name = "RowNum",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = headerText2,
                Width = 250,
                DataPropertyName = "Unit2",
                Name = "Unit2",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = headerText3,
                Width = 250,
                DataPropertyName = "Unit3",
                Name = "Unit3",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = headerText4,
                Width = 250,
                DataPropertyName = "Unit4",
                Name = "Unit4",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
        }

        private void CreateDgvColumnsForTable8(DataGridView dgvReport)
        {
            dgvReport.AllowUserToAddRows = true;
            dgvReport.ColumnHeadersVisible = true;
            var headerText1 =
                "№ п/п";
            var headerText2 =
                "Куда направлено";
            var headerText3 =
                "Реквизиты документа (при наличии)";
            var headerText4 =
                "Краткое содержание (суть предложений)";

            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = headerText1,
                Width = 50,
                DataPropertyName = "RowNum",
                Name = "RowNum",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = headerText2,
                Width = 250,
                DataPropertyName = "Unit2",
                Name = "Unit2",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = headerText3,
                Width = 250,
                DataPropertyName = "Unit3",
                Name = "Unit3",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = headerText4,
                Width = 250,
                DataPropertyName = "Unit4",
                Name = "Unit4",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
        }

        private void CreateDgvColumnsForTable2(DataGridView dgvReport)
        {
            dgvReport.AllowUserToAddRows = false;
            dgvReport.ColumnHeadersVisible = true;
            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Количество обращений",
                Width = 300,
                DataPropertyName = "Indicator",
                Name = "Indicator",
                ReadOnly = true,
                DefaultCellStyle =
                    new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "№ п/п",
                Width = 100,
                DataPropertyName = "RowNum",
                Name = "RowNum",
                ReadOnly = true,
                DefaultCellStyle =
                new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "За отчетный период (месяц)",
                Width = 100,
                DataPropertyName = "Unit2",
                Name = "Unit2",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Всего, с начала календарного года",
                Width = 100,
                DataPropertyName = "Unit3",
                Name = "Unit3",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Всего, с даты заключения соглашения",
                Width = 100,
                DataPropertyName = "Unit4",
                Name = "Unit4",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);


        }

        private void CreateDgvColumnsForTable3(DataGridView dgvReport)
        {
            dgvReport.AllowUserToAddRows = false;
            dgvReport.ColumnHeadersVisible = true;
            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Количество обратившихся за ИИС (завершенные обращения)",
                Width = 300,
                DataPropertyName = "Indicator",
                Name = "Indicator",
                ReadOnly = true,
                DefaultCellStyle =
                    new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "№ п/п",
                Width = 100,
                DataPropertyName = "RowNum",
                Name = "RowNum",
                ReadOnly = true,
                DefaultCellStyle =
                new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "За отчетный период (месяц)",
                Width = 100,
                DataPropertyName = "Unit2",
                Name = "Unit2",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Всего, с начала календарного года",
                Width = 100,
                DataPropertyName = "Unit3",
                Name = "Unit3",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Всего, с даты заключения соглашения",
                Width = 100,
                DataPropertyName = "Unit4",
                Name = "Unit4",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
        }

        private void CreateDgvColumnsForTable4(DataGridView dgvReport)
        {
            dgvReport.AllowUserToAddRows = false;
            dgvReport.ColumnHeadersVisible = true;
            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Количество индивидуально проинформированных застрахованных лиц",
                Width = 300,
                DataPropertyName = "Indicator",
                Name = "Indicator",
                ReadOnly = true,
                DefaultCellStyle =
                    new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "№ п/п",
                Width = 100,
                DataPropertyName = "RowNum",
                Name = "RowNum",
                ReadOnly = true,
                DefaultCellStyle =
                new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "За отчетный период (месяц)",
                Width = 100,
                DataPropertyName = "Unit2",
                Name = "Unit2",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Всего, с начала календарного года",
                Width = 100,
                DataPropertyName = "Unit3",
                Name = "Unit3",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Всего, с даты заключения соглашения",
                Width = 100,
                DataPropertyName = "Unit4",
                Name = "Unit4",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
        }


        private void CreateDgvColumnsForTable6(DataGridView dgvReport)
        {
            dgvReport.AllowUserToAddRows = false;
            dgvReport.ColumnHeadersVisible = true;
            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Количество завершённых рассмотрением обоснованных жалоб, по которым проведена ЭКМП",
                Width = 300,
                DataPropertyName = "Indicator",
                Name = "Indicator",
                ReadOnly = true,
                DefaultCellStyle =
                    new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "№ п/п",
                Width = 100,
                DataPropertyName = "RowNum",
                Name = "RowNum",
                ReadOnly = true,
                DefaultCellStyle =
                new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "За отчетный период (месяц)",
                Width = 100,
                DataPropertyName = "Unit2",
                Name = "Unit2",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Всего, с начала календарного года",
                Width = 100,
                DataPropertyName = "Unit3",
                Name = "Unit3",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Всего, с даты заключения соглашения",
                Width = 100,
                DataPropertyName = "Unit4",
                Name = "Unit4",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
        }


        private void CreateDgvColumnsForTable159(DataGridView dgvReport)
        {
            dgvReport.AllowUserToAddRows = false;
            dgvReport.ColumnHeadersVisible = false;
            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Информация",
                Width = 1000,
                DataPropertyName = "Indicator",
                Name = "Indicator",
            };
            dgvReport.Columns.Add(column);
        }


        private void CreateDgvColumnsForTable3_1(DataGridView dgvReport)
        {
            dgvReport.AllowUserToAddRows = true;
            dgvReport.ColumnHeadersVisible = true;
            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Тема (предмет) обращения",
                Width = 400,
                DataPropertyName = "Indicator",
                Name = "Indicator",
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Количество обратившихся участников СВО",
                Width = 200,
                DataPropertyName = "Unit2",
                Name = "Unit2",
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Количество обратившихся членов семьи",
                Width = 200,
                DataPropertyName = "Unit3",
                Name = "Unit3",
            };
            dgvReport.Columns.Add(column);
        }


        private void CreateDgvColumnsForTable4_1(DataGridView dgvReport)
        {
            dgvReport.AllowUserToAddRows = true;
            dgvReport.ColumnHeadersVisible = true;
            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Повод информирования",
                Width = 400,
                DataPropertyName = "Indicator",
                Name = "Indicator",
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Количество обратившихся участников СВО",
                Width = 200,
                DataPropertyName = "Unit2",
                Name = "Unit2",
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Количество обратившихся членов семьи",
                Width = 200,
                DataPropertyName = "Unit3",
                Name = "Unit3",
            };
            dgvReport.Columns.Add(column);
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


        public void SetFormula()
        {
            if (GetCurrentTheme() == "Таблица 2")
            {
                foreach (var row in _rows.Reverse())
                {
                    if (row.Key == "1.2.")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "1.2.а)" || x.Key == "1.2.б)" || x.Key == "1.2.в)").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "1.2.а)" || x.Key == "1.2.б)" || x.Key == "1.2.в)").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        row.Value.Cells[4].Value = _rows.Where(x => x.Key == "1.2.а)" || x.Key == "1.2.б)" || x.Key == "1.2.в)").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[4].Value));
                        continue;
                    }

                    if (row.Key == "1.1.")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "1.1.а)" || x.Key == "1.1.б)" || x.Key == "1.1.в)").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "1.1.а)" || x.Key == "1.1.б)" || x.Key == "1.1.в)").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        row.Value.Cells[4].Value = _rows.Where(x => x.Key == "1.1.а)" || x.Key == "1.1.б)" || x.Key == "1.1.в)").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[4].Value));
                        continue;
                    }

                    if (row.Key == "1.а)")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "1.1.а)" || x.Key == "1.2.а)").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "1.1.а)" || x.Key == "1.2.а)").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        row.Value.Cells[4].Value = _rows.Where(x => x.Key == "1.1.а)" || x.Key == "1.2.а)").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[4].Value));
                        continue;
                    }

                    if (row.Key == "1.б)")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "1.1.б)" || x.Key == "1.2.б)").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "1.1.б)" || x.Key == "1.2.б)").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        row.Value.Cells[4].Value = _rows.Where(x => x.Key == "1.1.б)" || x.Key == "1.2.б)").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[4].Value));
                        continue;
                    }

                    if (row.Key == "1.в)")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "1.1.в)" || x.Key == "1.2.в)").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "1.1.в)" || x.Key == "1.2.в)").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        row.Value.Cells[4].Value = _rows.Where(x => x.Key == "1.1.в)" || x.Key == "1.2.в)").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[4].Value));
                        continue;
                    }

                    if (row.Key == "1.")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "1.1." || x.Key == "1.2.").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "1.1." || x.Key == "1.2.").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        row.Value.Cells[4].Value = _rows.Where(x => x.Key == "1.1." || x.Key == "1.2.").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[4].Value));
                        continue;
                    }
                }
            }
            else if (GetCurrentTheme() == "Таблица 3" || GetCurrentTheme() == "Таблица 4" || GetCurrentTheme() == "Таблица 6")
            {
                foreach (var row in _rows.Reverse())
                    if (row.Key == "1.")
                    {
                        row.Value.Cells[2].Value = _rows.Where(x => x.Key == "1.1" || x.Key == "1.2").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[2].Value));
                        row.Value.Cells[3].Value = _rows.Where(x => x.Key == "1.1" || x.Key == "1.2").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[3].Value));
                        row.Value.Cells[4].Value = _rows.Where(x => x.Key == "1.1" || x.Key == "1.2").Sum(x => GlobalUtils.TryParseDecimal(x.Value.Cells[4].Value));
                        continue;
                    }

            }

        }

        private void FillDgvForms159(DataGridView dgvReport, string form)
        {
            var reportDoffDto = Report.ReportDataList?.Single(x => x.Theme == form);
            if (reportDoffDto?.Data == null || reportDoffDto.Data.Length == 0)
            {
                return;
            }

            var rows = ThemeTextData.Tables_fromxml.Where(x => x.TableName_fromxml == form).SelectMany(x => x.Rows_fromxml).ToList();
            foreach (DataGridViewRow row in dgvReport.Rows)
            {

                var data = reportDoffDto.Data.SingleOrDefault();
                if (data != null)
                {
                    row.Cells[0].Value = data.Column1.ToString();

                }
            }
        }

        private void FillDgvForms2(DataGridView dgvReport, string form)
        {
            var reportDoffDto = Report.ReportDataList?.Single(x => x.Theme == form);
            if (reportDoffDto?.Data == null || reportDoffDto.Data.Length == 0)
            {
                return;
            }

            var rows = ThemeTextData.Tables_fromxml.Where(x => x.TableName_fromxml == form).SelectMany(x => x.Rows_fromxml).ToList();
            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                var rowNum = row.Cells[1].Value.ToString().Trim();
                bool exclusionsRow = rows.Single(x => x.RowNum_fromxml == rowNum).Exclusion_fromxml;

                var data = reportDoffDto.Data.SingleOrDefault(x => x.RowNum == rowNum);
                if (data != null)
                {
                    row.Cells[2].Value = data.Column1.ToString();
                }


                var yearThemeData = Client.GetDoffYearData(new GetDoffYearDataRequest(new GetDoffYearDataRequestBody
                {
                    fillial = FilialCode,
                    theme = form,
                    yymm = Report.Yymm,
                    rowNum = rowNum
                })).Body.GetDoffYearDataResult;
                if (yearThemeData != null)
                {
                    row.Cells[3].Value = yearThemeData.Column2;

                }

                var beginningThemeData = Client.GetDoffBeginningData(new GetDoffBeginningDataRequest(new GetDoffBeginningDataRequestBody
                {
                    fillial = FilialCode,
                    theme = form,
                    yymm = Report.Yymm,
                    rowNum = rowNum
                })).Body.GetDoffBeginningDataResult;
                if (beginningThemeData != null)
                {
                    row.Cells[4].Value = beginningThemeData.Column3;

                }
            }

            SetFormula();
        }


        private void FillDgvForms346(DataGridView dgvReport, string form)
        {
            var reportDoffDto = Report.ReportDataList?.Single(x => x.Theme == form);
            if (reportDoffDto?.Data == null || reportDoffDto.Data.Length == 0)
            {
                return;
            }

            var rows = ThemeTextData.Tables_fromxml.Where(x => x.TableName_fromxml == form).SelectMany(x => x.Rows_fromxml).ToList();
            foreach (DataGridViewRow row in dgvReport.Rows)
            {
                var rowNum = row.Cells[1].Value.ToString().Trim();
                bool exclusionsRow = rows.Single(x => x.RowNum_fromxml == rowNum).Exclusion_fromxml;

                var data = reportDoffDto.Data.SingleOrDefault(x => x.RowNum == rowNum);
                if (data != null)
                {
                    row.Cells[2].Value = data.Column1.ToString();
                }


                var yearThemeData = Client.GetDoffYearData(new GetDoffYearDataRequest(new GetDoffYearDataRequestBody
                {
                    fillial = FilialCode,
                    theme = form,
                    yymm = Report.Yymm,
                    rowNum = rowNum
                })).Body.GetDoffYearDataResult;
                if (yearThemeData != null)
                {
                    row.Cells[3].Value = yearThemeData.Column2;

                }

                var beginningThemeData = Client.GetDoffBeginningData(new GetDoffBeginningDataRequest(new GetDoffBeginningDataRequestBody
                {
                    fillial = FilialCode,
                    theme = form,
                    yymm = Report.Yymm,
                    rowNum = rowNum
                })).Body.GetDoffBeginningDataResult;
                if (beginningThemeData != null)
                {
                    row.Cells[4].Value = beginningThemeData.Column3;

                }
            }

            SetFormula();
        }


        private void FillThemesForms159(DataGridView dgvReport, string form)
        {
            var reportDto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportDto == null)
            {
                return;
            }
            {
                var dataList = new List<ReportDoffDataDto>();

                foreach (DataGridViewRow row in Dgv.Rows)
                {
                    var data = new ReportDoffDataDto
                    {
                        Column1 = row.Cells[0].Value.ToString()
                    };
                    dataList.Add(data);
                }

                reportDto.Data = dataList.ToArray();
            }

        }

        private void FillThemesForms2(DataGridView dgvReport, string form)
        {
            var reportDoffDto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportDoffDto == null)
            {
                return;
            }

            reportDoffDto.Data = (from DataGridViewRow row in dgvReport.Rows
                                  let rowNum = row.Cells[1].Value.ToString().Trim()
                                  where !IsNotNeedFillRow(form, rowNum)
                                  select new ReportDoffDataDto
                                  {
                                      RowNum = rowNum,
                                      Column1 = row.Cells[2].Value.ToString(),
                                  }).ToArray();
        }



        private void FillThemesForms346(DataGridView dgvReport, string form)
        {
            var reportDoffDto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportDoffDto == null)
            {
                return;
            }

            reportDoffDto.Data = (from DataGridViewRow row in dgvReport.Rows
                                  let rowNum = row.Cells[1].Value.ToString().Trim()
                                  where !IsNotNeedFillRow(form, rowNum)
                                  select new ReportDoffDataDto
                                  {
                                      RowNum = rowNum,
                                      Column1 = row.Cells[2].Value.ToString(),
                                  }).ToArray();
        }
    }
}