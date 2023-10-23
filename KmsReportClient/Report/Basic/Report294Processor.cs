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
    public class Report294Processor : AbstractReportProcessor<Report294>
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        private readonly List<string> disiaseColumnsText = new List<string> {
            "Онкологические заболевания",
            "Заболевания эндокринной системы",
            "Бронхолегочные заболевания",
            "Болезни системы кровообращения",
            "Прочие неинфекционные заболевания"
        };

        private readonly List<string> dispColumnsText = new List<string> {
            "SMS рассылка",
            "Почтовые рассылки",
            "Телефонный обзвон",
            "Мессенджеры",
            "Электронная почта",
            "Адресный обход",
            "Иные способы"
        };

        private readonly string[] forms1 = {
            "Таблица 1", "Таблица 2", "Таблица 7", "Таблица 8", "Таблица 9", "Эффективность"
        };

        private readonly string[] forms2 = { "Таблица 3", "Таблица 4", "Таблица 5" };
        private readonly string[] forms3 = { "Таблица 6" };

        private readonly List<string> singleColumnsList = new List<string> { "За отчетный период" };

        public Report294Processor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
            base(inClient, dgv, cmb, txtb, page,
                XmlFormTemplate.F294.GetDescription(),
                Log,
                ReportGlobalConst.Report294,
                reportsDictionary)
        {
            InitReport();
        }

        public override void InitReport()
        {
            Report = new Report294 { ReportDataList = new Report294Dto[ThemesList.Count], IdType = IdReportType };
            int i = 0;
            foreach (var theme in ThemesList.Select(x => x.Key))
            {
                Report.ReportDataList[i++] = new Report294Dto { Theme = theme };
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
                    reportType = ReportType.F294
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as Report294;
        }

        public override void MapForAutoFill(AbstractReport report)
        {
            if (report == null)
            {
                return;
            }
            var inReport = report as Report294;

            var index = Report.ReportDataList.ToList().FindIndex(x => x.Theme == Cmb.Text);
            var inTheme = inReport.ReportDataList.Single(x => x.Theme == Cmb.Text);
            Report.ReportDataList[index] = inTheme;
        }
        public override void SaveReportDataSourceExcel()
        { }
        public override void SaveReportDataSourceHandle()
        { }
        public override void FillDataGridView(string form)
        {
            if (forms1.Contains(form))
            {
                FillDgwForms1(Dgv, form);
            }
            else if (forms2.Contains(form))
            {
                FillDgwForms2(Dgv, form);
            }
            else
            {
                FillDgwForms3(Dgv, form);
            }
        }

        protected override void FillReport(string form)
        {
            if (form == null)
            {
                return;
            }

            if (forms1.Contains(form))
            {
                FillThemesForms1(Dgv, form);
            }
            else if (forms2.Contains(form))
            {
                FillThemesForms2(Dgv, form);
            }
            else
            {
                FillThemesForms3(Dgv, form);
            }
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
                    reportType = ReportType.F294
                }
            };
            var response = Client.CollectSummaryReport(request);
            Report = response.Body.CollectSummaryReportResult as Report294;
            Report.IdType = IdReportType;
            Report.Yymm = yymmEnd;
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
                    reportType = ReportType.F294
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as Report294;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
        }

        public override void ToExcel(string filename, string filialName)
        {
            var excel = new ExcelF294Creator(filename, ExcelForm.F294, Report.Yymm, filialName);
            var yearReport = FillYearReport();
            excel.CreateReport(Report, yearReport);
        }

        public override bool IsVisibleBtnDownloadExcel() => false;

        public override bool IsVisibleBtnHandle() => false;

        public override bool IsVisibleBtnSummary() => false;

        public override string ValidReport()
        {
            var message = "";
            foreach (var data in Report.ReportDataList)
            {
                if (data.Data == null)
                {
                    continue;
                }

                if (data.Theme == "Таблица 8")
                {
                    int sum11 = 0;
                    int sum12 = 0;
                    int sum13 = 0;
                    int sum14 = 0;
                    int sum20 = 0;
                    int sum21 = 0;
                    int sum30 = 0;
                    int sum31 = 0;
                    foreach (var table in data.Data)
                    {
                        switch (table.RowNum)
                        {
                            case "01.1":
                                sum11 = table.CountPpl;
                                break;
                            case "01.2":
                                sum12 = table.CountPpl;
                                break;
                            case "01.3":
                                sum13 = table.CountPpl;
                                break;
                            case "01.4":
                                sum14 = table.CountPpl;
                                break;
                            case "02":
                                sum20 = table.CountPpl;
                                break;
                            case "02.1":
                                sum21 = table.CountPpl;
                                break;
                            case "03":
                                sum30 = table.CountPpl;
                                break;
                            case "03.1":
                                sum31 = table.CountPpl;
                                break;
                        }
                    }

                    string localMessage = "";
                    if (sum11 < sum12)
                    {
                        localMessage += "Строка 01.1 должна быть больше или равна 01.2" + Environment.NewLine;
                    }

                    if (sum13 < sum14)
                    {
                        localMessage += "Строка 01.3 должна быть больше или равна 01.4" + Environment.NewLine;
                    }

                    if (sum20 < sum21)
                    {
                        localMessage += "Строка 02 должна быть больше или равна 02.1" + Environment.NewLine;
                    }

                    if (sum30 < sum31)
                    {
                        localMessage += "Строка 03 должна быть больше или равна 03.1" + Environment.NewLine;
                    }

                    if (localMessage.Length > 0)
                    {
                        message = "Форма 294. Таблица 8. " + Environment.NewLine + localMessage;
                    }
                }
                else if (data.Theme == "Таблица 9")
                {
                    int sum10 = 0;
                    int sum11 = 0;
                    int sum24 = 0;
                    int sum30 = 0;

                    foreach (var table in data.Data)
                    {
                        switch (table.RowNum)
                        {
                            case "01":
                                sum10 = table.CountPpl;
                                break;
                            case "01.1":
                                sum11 = table.CountPpl;
                                break;
                            case "02.4":
                                sum24 = table.CountPpl;
                                break;
                            case "03":
                                sum30 = table.CountPpl;
                                break;
                        }
                    }

                    string localMessage = "";
                    if (sum10 < sum11)
                    {
                        localMessage += "Строка 01 должна быть больше или равна 01.1" + Environment.NewLine;
                    }

                    if (sum24 > sum30)
                    {
                        localMessage += "Строка 03 должна быть меньше или равна 02.4" + Environment.NewLine;
                    }

                    if (localMessage.Length > 0)
                    {
                        message = "Форма 294. Таблица 9. " + Environment.NewLine + localMessage;
                    }
                }
            }

            return message;
        }

        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            if (forms1.Contains(form))
            {
                CreateDgvColumnsForTheme(Dgv, 400, singleColumnsList);
            }
            else if (forms2.Contains(form))
            {
                CreateDgvColumnsForTheme(Dgv, 400, dispColumnsText);
            }
            else if (forms3.Contains(form))
            {
                CreateDgvColumnsForTheme(Dgv, 400, disiaseColumnsText);
            }

            int countRows = ThemeTextData.Tables_fromxml.Single(x => x.TableName_fromxml == form).RowsCount_fromxml;
            foreach (var row in table)
            {
                var dgvRow = new DataGridViewRow();
                var cellName = new DataGridViewTextBoxCell { Value = row.RowText_fromxml };
                var cellNum = new DataGridViewTextBoxCell { Value = row.RowNum_fromxml };
                var cellPpl = new DataGridViewTextBoxCell { Value = "человек" };
                dgvRow.Cells.Add(cellName);
                dgvRow.Cells.Add(cellNum);
                dgvRow.Cells.Add(cellPpl);

                for (int i = 3; i < countRows; i++)
                {
                    var cell = new DataGridViewTextBoxCell { Value = row.Exclusion_fromxml ? "x" : "0" };
                    dgvRow.Cells.Add(cell);
                }

                int rowIndex = Dgv.Rows.Add(dgvRow);
                if (row.Exclusion_fromxml)
                {
                    Dgv.Rows[rowIndex].ReadOnly = true;
                    Dgv.Rows[rowIndex].DefaultCellStyle.BackColor = Color.LightCyan;
                }
            }
        }

        private Report294 FillYearReport()
        {
            var request = new CollectSummaryReportRequest
            {
                Body = new CollectSummaryReportRequestBody
                {
                    filials = new ArrayOfString { FilialCode },
                    status = ReportStatus.Saved,
                    yymmStart = Report.Yymm.Substring(0, 2) + "01",
                    yymmEnd = Report.Yymm,
                    reportType = ReportType.F294
                }
            };

            return Client.CollectSummaryReport(request).Body.CollectSummaryReportResult as Report294;
        }

        private void FillDgwForms1(DataGridView dgvReport, string form)
        {
            var report294Dto = Report.ReportDataList.FirstOrDefault(x => x.Theme == form);
            if (report294Dto == null)
            {
                return;
            }

            if (report294Dto.Data != null && report294Dto.Data.Length > 0)
            {
                foreach (DataGridViewRow row in dgvReport.Rows)
                {
                    var rowNum = row.Cells[1].Value.ToString();
                    var data = report294Dto.Data.SingleOrDefault(x => x.RowNum == rowNum);
                    if (data != null)
                    {
                        row.Cells[3].Value = data.CountPpl;
                    }
                }
            }
        }

        private void FillDgwForms2(DataGridView dgvReport, string form)
        {
            var report294Dto = Report.ReportDataList.FirstOrDefault(x => x.Theme == form);
            if (report294Dto == null)
            {
                return;
            }
            if (report294Dto.Data != null && report294Dto.Data.Length > 0)
            {
                foreach (DataGridViewRow row in dgvReport.Rows)
                {
                    var rowNum = row.Cells[1].Value.ToString();
                    var data = report294Dto.Data.SingleOrDefault(x => x.RowNum == rowNum);
                    if (data != null)
                    {
                        row.Cells[3].Value = data.CountSms;
                        row.Cells[4].Value = data.CountPost;
                        row.Cells[5].Value = data.CountPhone;
                        row.Cells[6].Value = data.CountMessengers;
                        row.Cells[7].Value = data.CountEmail;
                        row.Cells[8].Value = data.CountAddress;
                        row.Cells[9].Value = data.CountAnother;
                    }
                }
            }
        }

        private void FillDgwForms3(DataGridView dgvReport, string form)
        {
            var report294Dto = Report.ReportDataList.Single(x => x.Theme == form);
            if (report294Dto.Data != null && report294Dto.Data.Length > 0)
            {
                foreach (DataGridViewRow row in dgvReport.Rows)
                {
                    var rowNum = row.Cells[1].Value.ToString();
                    var data = report294Dto.Data.SingleOrDefault(x => x.RowNum == rowNum);
                    if (data != null)
                    {
                        row.Cells[3].Value = data.CountOncologicalDisease;
                        row.Cells[4].Value = data.CountEndocrineDisease;
                        row.Cells[5].Value = data.CountBronchoDisease;
                        row.Cells[6].Value = data.CountBloodDisease;
                        row.Cells[7].Value = data.CountAnotherDisease;
                    }
                }
            }
        }

        private void FillThemesForms1(DataGridView dgvReport, string form)
        {
            var report294Dto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (report294Dto != null)
            {
                report294Dto.Data = (from DataGridViewRow row in dgvReport.Rows
                                     let rowNum = row.Cells[1].Value.ToString()
                                     where !IsNotNeedFillRow(form, rowNum)
                                     select new Report294DataDto
                                     {
                                         RowNum = rowNum,
                                         CountPpl = GlobalUtils.TryParseInt(row.Cells[3].Value)
                                     }).ToArray();
            }
        }

        private void FillThemesForms2(DataGridView dgvReport, string form)
        {
            var report294Dto = Report.ReportDataList.FirstOrDefault(x => x.Theme == form);
            if (report294Dto != null)
            {
                report294Dto.Data = (from DataGridViewRow row in dgvReport.Rows
                                     let rowNum = row.Cells[1].Value.ToString()
                                     where !IsNotNeedFillRow(form, rowNum)
                                     select new Report294DataDto
                                     {
                                         RowNum = rowNum,
                                         CountSms = GlobalUtils.TryParseInt(row.Cells[3].Value),
                                         CountPost = GlobalUtils.TryParseInt(row.Cells[4].Value),
                                         CountPhone = GlobalUtils.TryParseInt(row.Cells[5].Value),
                                         CountMessengers = GlobalUtils.TryParseInt(row.Cells[6].Value),
                                         CountEmail = GlobalUtils.TryParseInt(row.Cells[7].Value),
                                         CountAddress = GlobalUtils.TryParseInt(row.Cells[8].Value),
                                         CountAnother = GlobalUtils.TryParseInt(row.Cells[9].Value)
                                     }).ToArray();
            }
        }

        private void FillThemesForms3(DataGridView dgvReport, string form)
        {
            var report294Dto = Report.ReportDataList.FirstOrDefault(x => x.Theme == form);
            if (report294Dto != null)
            {
                var dataList = new List<Report294DataDto>();

                foreach (DataGridViewRow row in dgvReport.Rows)
                {
                    var rowNum = row.Cells[1].Value.ToString();
                    if (IsNotNeedFillRow(form, rowNum))
                    {
                        continue;
                    }

                    var data = new Report294DataDto
                    {
                        RowNum = rowNum,
                        CountOncologicalDisease = GlobalUtils.TryParseInt(row.Cells[3].Value),
                        CountEndocrineDisease = GlobalUtils.TryParseInt(row.Cells[4].Value),
                        CountBronchoDisease = GlobalUtils.TryParseInt(row.Cells[5].Value),
                        CountBloodDisease = GlobalUtils.TryParseInt(row.Cells[6].Value),
                        CountAnotherDisease = GlobalUtils.TryParseInt(row.Cells[7].Value)
                    };
                    dataList.Add(data);
                }

                report294Dto.Data = dataList.ToArray();
            }
        }

        private void CreateDgvColumnsForTheme(DataGridView dgvReport, int widthFirstColumn, List<string> columns)
        {
            CreateDgvCommonColumns(dgvReport, widthFirstColumn);
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

        private void CreateDgvCommonColumns(DataGridView dgvReport, int widthFirstColumn)
        {
            dgvReport.AllowUserToAddRows = false;
            dgvReport.ColumnHeadersVisible = true;
            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Наименование показателя",
                Width = widthFirstColumn,
                DataPropertyName = "Indicator",
                Name = "Indicator",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Azure }
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
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = "Единица измерения",
                Width = 80,
                DataPropertyName = "Unit",
                Name = "Unit",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
        }
    }
}