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
    internal class Report262Processor : AbstractReportProcessor<Report262>
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        public Report262Processor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
            base(inClient, dgv, cmb, txtb, page,
                XmlFormTemplate.F262.GetDescription(),
                Log,
                ReportGlobalConst.Report262,
                reportsDictionary)
        {
            InitReport();
        }

        public override void InitReport()
        {
            Report = new Report262 { ReportDataList = new Report262Dto[ThemesList.Count], IdType = IdReportType };
            int i = 0;
            foreach (var theme in ThemesList.Select(x => x.Key))
            {
                Report.ReportDataList[i++] = new Report262Dto { Theme = theme };
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
                    reportType = ReportType.F262
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as Report262;
        }

        public override void FillDataGridView(string form)
        {
            Dgv.RowHeadersWidth = 30;
            if (form == null)
            {
                return;
            }

            var reportDto = Report.ReportDataList.Single(x => x.Theme == form);

            if (form == "Таблица 1")
            {
                if (reportDto.Data == null || reportDto.Data.Length == 0)
                {
                    return;
                }

                foreach (DataGridViewRow row in Dgv.Rows)
                {
                    var code = GlobalUtils.TryParseInt(row.Cells[1].Value);
                    var data =
                        reportDto.Data.SingleOrDefault(
                            x => x.RowNum == code);
                    if (data != null)
                    {
                        row.Cells[3].Value = data.CountPpl;
                    }
                }
            }
            else if (form == "Таблица 2")
            {
                if (reportDto.Data == null || reportDto.Data.Length == 0)
                {
                    return;
                }

                foreach (DataGridViewRow row in Dgv.Rows)
                {
                    var data = reportDto.Data[0];
                    if (data == null)
                    {
                        continue;
                    }

                    row.Cells[2].Value = data.CountSms;
                    row.Cells[3].Value = data.CountPost;
                    row.Cells[4].Value = data.CountPhone;
                    row.Cells[5].Value = data.CountMessengers;
                    row.Cells[6].Value = data.CountEmail;
                    row.Cells[7].Value = data.CountAddress;
                    row.Cells[8].Value = data.CountAnother;
                }
            }
            else if (form == "Таблица 3")
            {
                Dgv.Rows.Clear();
                if (reportDto.Table3 == null || reportDto.Table3.Length == 0)
                {
                    return;
                }

                foreach (var data in reportDto.Table3)
                {
                    Dgv.Rows.Add(data.Mo, data.CountUnit,
                        data.CountUnitChild, data.CountUnitWithSp,
                        data.CountUnitWithSpChild, data.CountChannelSp,
                        data.CountChannelSpChild,
                        data.CountChannelPhone, data.CountChannelPhoneChild,
                        data.CountChannelTerminal,
                        data.CountChannelTerminalChild,
                        data.CountChannelAnother,
                        data.CountChannelAnotherChild);
                }

                if (CurrentUser.IsMain)
                {
                    var total = new Report262Table3Data
                    {
                        Mo = "Итого:",
                        CountUnit = reportDto.Table3.Sum(x => x.CountUnit),
                        CountUnitChild =
                            reportDto.Table3.Sum(x => x.CountUnitChild),
                        CountUnitWithSp =
                            reportDto.Table3.Sum(x => x.CountUnitWithSp),
                        CountUnitWithSpChild =
                            reportDto.Table3.Sum(x => x.CountUnitWithSpChild),
                        CountChannelSp =
                            reportDto.Table3.Sum(x => x.CountChannelSp),
                        CountChannelSpChild =
                            reportDto.Table3.Sum(x => x.CountChannelSpChild),
                        CountChannelPhone =
                            reportDto.Table3.Sum(x => x.CountChannelPhone),
                        CountChannelPhoneChild =
                            reportDto.Table3.Sum(x => x.CountChannelPhoneChild),
                        CountChannelTerminal =
                            reportDto.Table3.Sum(x => x.CountChannelTerminal),
                        CountChannelTerminalChild =
                            reportDto.Table3.Sum(x =>
                                x.CountChannelTerminalChild),
                        CountChannelAnother =
                            reportDto.Table3.Sum(x => x.CountChannelAnother),
                        CountChannelAnotherChild =
                            reportDto.Table3.Sum(
                                x => x.CountChannelAnotherChild)
                    };
                    Dgv.Rows.Add(total.Mo, total.CountUnit,
                        total.CountUnitChild, total.CountUnitWithSp,
                        total.CountUnitWithSpChild, total.CountChannelSp,
                        total.CountChannelSpChild,
                        total.CountChannelPhone, total.CountChannelPhoneChild,
                        total.CountChannelTerminal,
                        total.CountChannelTerminalChild,
                        total.CountChannelAnother,
                        total.CountChannelAnotherChild);
                    Dgv.AllowUserToAddRows = false;
                    int lastRow = Dgv.Rows.Count - 1;
                    Dgv.Rows[lastRow].DefaultCellStyle =
                        new DataGridViewCellStyle { BackColor = Color.Azure };
                }
            }
        }

        protected override void FillReport(string form)
        {
            if (form == null)
            {
                return;
            }

            var reportDto = Report.ReportDataList.Single(x => x.Theme == form);

            if (form == "Таблица 1")
            {
                var dataList = new List<Report262DataDto>();

                foreach (DataGridViewRow row in Dgv.Rows)
                {
                    var data = new Report262DataDto
                    {
                        RowNum = GlobalUtils.TryParseInt(row.Cells[1].Value),
                        CountPpl = GlobalUtils.TryParseInt(row.Cells[3].Value)
                    };
                    dataList.Add(data);
                }

                reportDto.Data = dataList.ToArray();
            }
            else if (form == "Таблица 2")
            {
                var dataList = new List<Report262DataDto>();

                foreach (DataGridViewRow row in Dgv.Rows)
                {
                    var data = new Report262DataDto
                    {
                        RowNum = 0,
                        CountSms = GlobalUtils.TryParseInt(row.Cells[2].Value),
                        CountPost = GlobalUtils.TryParseInt(row.Cells[3].Value),
                        CountPhone =
                            GlobalUtils.TryParseInt(row.Cells[4].Value),
                        CountMessengers =
                            GlobalUtils.TryParseInt(row.Cells[5].Value),
                        CountEmail =
                            GlobalUtils.TryParseInt(row.Cells[6].Value),
                        CountAddress =
                            GlobalUtils.TryParseInt(row.Cells[7].Value),
                        CountAnother =
                            GlobalUtils.TryParseInt(row.Cells[8].Value)
                    };
                    dataList.Add(data);
                }

                reportDto.Data = dataList.ToArray();
            }
            else
            {
                var table3DataList = new List<Report262Table3Data>();

                foreach (DataGridViewRow row in Dgv.Rows)
                {
                    string mo = row.Cells[0].Value?.ToString() ?? "";
                    if (string.IsNullOrEmpty(mo))
                    {
                        continue;
                    }

                    var data = new Report262Table3Data
                    {
                        Mo = mo.Replace("  ", " ").Replace("\n", ""),
                        CountUnit = GlobalUtils.TryParseInt(row.Cells[1].Value),
                        CountUnitChild =
                            GlobalUtils.TryParseInt(row.Cells[2].Value),
                        CountUnitWithSp =
                            GlobalUtils.TryParseInt(row.Cells[3].Value),
                        CountUnitWithSpChild =
                            GlobalUtils.TryParseInt(row.Cells[4].Value),
                        CountChannelSp =
                            GlobalUtils.TryParseInt(row.Cells[5].Value),
                        CountChannelSpChild =
                            GlobalUtils.TryParseInt(row.Cells[6].Value),
                        CountChannelPhone =
                            GlobalUtils.TryParseInt(row.Cells[7].Value),
                        CountChannelPhoneChild =
                            GlobalUtils.TryParseInt(row.Cells[8].Value),
                        CountChannelTerminal =
                            GlobalUtils.TryParseInt(row.Cells[9].Value),
                        CountChannelTerminalChild =
                            GlobalUtils.TryParseInt(row.Cells[10].Value),
                        CountChannelAnother =
                            GlobalUtils.TryParseInt(row.Cells[11].Value),
                        CountChannelAnotherChild =
                            GlobalUtils.TryParseInt(row.Cells[12].Value)
                    };
                    table3DataList.Add(data);
                }

                reportDto.Table3 = table3DataList.ToArray();
            }
        }

        public override void ToExcel(string filename, string filialName)
        {
            var excel = new ExcelF262Creator(filename, ExcelForm.F262, Report.Yymm, filialName);
            var yearReport = FillYearReport();
            excel.CreateReport(Report, yearReport);
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
                    reportType = ReportType.F262
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as Report262;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
        }

        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status)
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
                    reportType = ReportType.F262
                }
            };
            Report = Client.CollectSummaryReport(request).Body.CollectSummaryReportResult as Report262;
            Report.IdType = IdReportType;
            Report.Yymm = yymmEnd;
        }

        public override bool IsVisibleBtnDownloadExcel() => Cmb.Text == "Таблица 3";

        public override string ValidReport()
        {
            var message = "";
            foreach (var data in Report.ReportDataList)
            {
                if (data.Theme == "Таблица 3")
                {
                    var i = 1;
                    if (data.Table3 == null || data.Table3.Length == 0)
                    {
                        message = @"Перед выгрузкой в Excel необходимо заполнить таблицу 3";
                        continue;
                    }

                    foreach (var table3 in data.Table3)
                    {
                        if (string.IsNullOrEmpty(table3.Mo))
                        {
                            message = $"Строка {i}. гр.1 не может быть пустым; ";
                        }
                        if (table3.CountUnit < 1)
                        {
                            message += $"Строка {i}. гр.2 должна быть больше 0" + Environment.NewLine;
                        }
                        if (table3.CountUnitChild > table3.CountUnit)
                        {
                            message += $"Строка {i}. гр.3 должно быть меньше или равна гр.2" + Environment.NewLine;
                        }
                        if (table3.CountUnitWithSpChild > table3.CountUnitWithSp)
                        {
                            message += $"Строка {i}. гр.5 должно быть меньше или равна гр.4" + Environment.NewLine;
                        }
                        if (table3.CountChannelSpChild > table3.CountChannelSp)
                        {
                            message += $"Строка {i}. гр.7 должно быть меньше или равна гр.6" + Environment.NewLine;
                        }
                        if (table3.CountChannelPhoneChild > table3.CountChannelPhone)
                        {
                            message += $"Строка {i}. гр.9 должно быть меньше или равна гр.8" + Environment.NewLine;
                        }
                        if (table3.CountChannelTerminalChild > table3.CountChannelTerminal)
                        {
                            message += $"Строка {i}. гр.11 должно быть меньше или равна гр.10" + Environment.NewLine;
                        }
                        if (table3.CountChannelAnotherChild > table3.CountChannelAnother)
                        {
                            message += $"Строка {i}. гр.13 должно быть меньше или равна гр.12" + Environment.NewLine;
                        }
                        if (table3.CountUnitWithSp > table3.CountUnit)
                        {
                            message += $"Строка {i}. гр.4 должно быть меньше или равна гр.2" + Environment.NewLine;
                        }
                        if (table3.CountUnitWithSpChild > table3.CountUnitChild)
                        {
                            message += $"Строка {i}. гр.5 должно быть меньше или равна гр.3" + Environment.NewLine;
                        }

                        var sumChild =
                            table3.CountChannelSpChild +
                            table3.CountChannelPhoneChild +
                            table3.CountChannelTerminalChild +
                            table3.CountChannelAnotherChild;
                        if (table3.CountUnitWithSpChild > sumChild)
                        {
                            message += $"Строка {i}. гр.5 не должна быть больше суммы гр.7, 9, 11, 13" + Environment.NewLine;
                        }

                        if (table3.CountUnitWithSpChild < 0)
                        {
                            message += $"Строка {i}. гр.5 не должна быть меньше 0" + Environment.NewLine;
                        }

                        var sum = table3.CountChannelSp +
                                  table3.CountChannelPhone +
                                  table3.CountChannelTerminal +
                                  table3.CountChannelAnother;
                        if (table3.CountUnitWithSp > sum)
                        {
                            message += $"Строка {i}. гр.4 не должна быть больше суммы гр. 6, 8, 10, 12" + Environment.NewLine;
                        }
                        if (table3.CountUnitWithSp < 0)
                        {
                            message += $"Строка {i}. гр.4 не должна быть меньше 0" + Environment.NewLine;
                        }

                        i++;
                    }

                    if (message.Length > 0)
                    {
                        message = "Форма 262. Таблица 3. " + Environment.NewLine + message;

                        Dgv.RowHeadersWidth = 60;
                        foreach (DataGridViewRow row in Dgv.Rows)
                        {
                            row.HeaderCell.Value = Convert.ToString(row.Index + 1);
                        }
                    }
                }
            }

            return message;
        }

        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            if (form == "Таблица 1")
            {
                CreateDgvColumnsForTable1(Dgv);
                foreach (var row in table)
                {
                    Dgv.Rows.Add(row.Name, row.Num, "человек", "0");
                }
            }
            else if (form == "Таблица 2")
            {
                CreateDgvColumnsForTable2(Dgv);
                foreach (var row in table)
                {
                    Dgv.Rows.Add(row.Name, "человек", "0", "0", "0", "0",
                        "0", "0", "0");
                }
            }
            else if (form == "Таблица 3")
            {
                Dgv.AllowUserToAddRows = true;
                CreateDgvColumnsForTable3(Dgv);
            }
        }

        private void CreateDgvColumnsForTable3(DataGridView dgvReport)
        {
            dgvReport.AllowUserToAddRows = true;
            var headerText1 =
                "Наименование МО, оказывающей в рамках ОМС первичную медико-санитарную помощь ";
            var headerText2 =
                "Количество МО, в том числе являющихся структурными подразделениями МО ";
            var headerText3 =
                "Количество МО, в т.ч. являющихся структурными подразделениями МО, " +
                "на базе которых функционируют каналы связи граждан с СП СМО ";
            var headerText4 =
                "посредством организации поста страхового представителя ";
            var headerText5 = "посредством прямой телефонной связи ";
            var headerText6 =
                "через терминал для связи со страховым представителем ";
            var headerText7 = "посредством иных каналов связи ";
            var childrenText = "в том числе детских ";

            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = headerText1 + Environment.NewLine + "(гр.1)",
                Width = 250,
                DataPropertyName = "Indicator",
                Name = "Indicator",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = headerText2 + Environment.NewLine + "(гр.2)",
                Width = 110,
                DataPropertyName = "Unit2",
                Name = "Unit2",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = childrenText + Environment.NewLine + "(гр.3)",
                Width = 55,
                DataPropertyName = "Unit3",
                Name = "Unit3",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = headerText3 + Environment.NewLine + "(гр.4)",
                Width = 140,
                DataPropertyName = "Unit4",
                Name = "Unit4",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = childrenText + Environment.NewLine + "(гр.5)",
                Width = 55,
                DataPropertyName = "Unit5",
                Name = "Unit5",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = headerText4 + Environment.NewLine + "(гр.6)",
                Width = 90,
                DataPropertyName = "Unit6",
                Name = "Unit6",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = childrenText + Environment.NewLine + "(гр.7)",
                Width = 55,
                DataPropertyName = "Unit7",
                Name = "Unit7",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = headerText5 + Environment.NewLine + "(гр.8)",
                Width = 80,
                DataPropertyName = "Unit8",
                Name = "Unit8",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = childrenText + Environment.NewLine + "(гр.9)",
                Width = 55,
                DataPropertyName = "Unit9",
                Name = "Unit9",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = headerText6 + Environment.NewLine + "(гр.10)",
                Width = 100,
                DataPropertyName = "Unit10",
                Name = "Unit10",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = childrenText + Environment.NewLine + "(гр.11)",
                Width = 55,
                DataPropertyName = "Unit11",
                Name = "Unit11",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = headerText7 + Environment.NewLine + "(гр.12)",
                Width = 80,
                DataPropertyName = "Unit12",
                Name = "Unit12",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = childrenText + Environment.NewLine + "(гр.13)",
                Width = 55,
                DataPropertyName = "Unit13",
                Name = "Unit13",
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            dgvReport.Columns.Add(column);
        }

        private void CreateDgvColumnsForTable2(DataGridView dgvReport)
        {
            dgvReport.AllowUserToAddRows = false;
            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Наименование показателя",
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
                HeaderText = "Единица измерения",
                Width = 70,
                DataPropertyName = "Unit",
                Name = "Unit",
                ReadOnly = true,
                DefaultCellStyle =
                    new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "СМС сообщения",
                Width = 70,
                DataPropertyName = "Sms",
                Name = "Sms"
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Почтовые рассылки",
                Width = 70,
                DataPropertyName = "Post",
                Name = "Post"
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "По телефону",
                Width = 70,
                DataPropertyName = "Phone",
                Name = "Phone"
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText =
                    "Системы обмена текстовыми сообщениями для мобильных платформ (мессенджеры)",
                Width = 90,
                DataPropertyName = "Messengers",
                Name = "Messengers"
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Электронная почта",
                Width = 80,
                DataPropertyName = "Email",
                Name = "Email"
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Адресный обход",
                Width = 70,
                DataPropertyName = "Address",
                Name = "Address"
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Иные способы индивидуального информирования",
                Width = 110,
                DataPropertyName = "Another",
                Name = "Another"
            };
            dgvReport.Columns.Add(column);
        }

        private void CreateDgvColumnsForTable1(DataGridView dgvReport)
        {
            dgvReport.AllowUserToAddRows = false;
            dgvReport.ColumnHeadersVisible = true;
            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Наименование показателя",
                Width = 400,
                DataPropertyName = "Indicator",
                Name = "Indicator",
                ReadOnly = true,
                DefaultCellStyle =
                    new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "№ строки",
                Width = 50,
                DataPropertyName = "NumRow",
                Name = "NumRow",
                ReadOnly = true,
                DefaultCellStyle =
                    new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Единица измерения",
                Width = 80,
                DataPropertyName = "Unit",
                Name = "Unit",
                ReadOnly = true,
                DefaultCellStyle =
                    new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "За отчетный период",
                Width = 100,
                DataPropertyName = "CountPeople",
                Name = "CountPeople"
            };
            dgvReport.Columns.Add(column);
        }


        private Report262 FillYearReport()
        {
            var request = new CollectSummaryReportRequest
            {
                Body = new CollectSummaryReportRequestBody
                {
                    filials = new ArrayOfString { FilialCode },
                    status = ReportStatus.Saved,
                    yymmStart = Report.Yymm.Substring(0, 2) + "01",
                    yymmEnd = Report.Yymm,
                    reportType = ReportType.F262
                }
            };
            return Client.CollectSummaryReport(request).Body.CollectSummaryReportResult as Report262;
        }

        public override void MapForAutoFill(AbstractReport report)
        {
            if (report == null)
            {
                return;
            }
            var inReport = report as Report262;

            var index = Report.ReportDataList.ToList().FindIndex(x => x.Theme == Cmb.Text);
            var inTheme = inReport.ReportDataList.Single(x => x.Theme == Cmb.Text);
            Report.ReportDataList[index] = inTheme;
        }
    }
}