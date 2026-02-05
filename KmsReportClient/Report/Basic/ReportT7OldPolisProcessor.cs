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
using KmsReportClient.Model.Excel;
using KmsReportClient.Model.XML;
using KmsReportClient.Support;
using NLog;

namespace KmsReportClient.Report.Basic
{
    internal class ReportT7OldPolisProcessor : AbstractReportProcessor<ReportT7OldPolis>
    {
        StackedHeaderDecorator DgvRender;

        private readonly List<string> based = new List<string>
        {
            "\r\nКол-во полисов старого образца  на 01.01.19 (данные постоянные)\r\n",
            "\r\nКол-во полисов старого образца  на 01.01 текущего года\r\n",
            "\r\nЧисленность на 01 число текущего месяца\r\n",
            "\r\nКол-во полисов старого образца  на 01 число текущего месяца\r\n",
            "\r\nДоля полисов старого образца   от численности на 01 число текущего месяца\r\n",
            "\r\nДинамика за отчетный год\r\n",
            "\r\nДинамика за период с 2019 на текущую отчетную дату\r\n"
        };

        private static readonly Logger Log = LogManager.GetCurrentClassLogger();
        public ReportT7OldPolisProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary,
                                        DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
        base(inClient, dgv, cmb, txtb, page,
            XmlFormTemplate.T7OldPolis.GetDescription(),
            Log,
            ReportGlobalConst.ReportT7OldPolis,
            reportsDictionary)
        {
            DgvRender = new StackedHeaderDecorator(Dgv);
            InitReport();
        }

        public override AbstractReport CollectReportFromWs(string yymm)
        {
            // Попытка загрузить существующий отчёт
            var request = new GetReportRequest
            {
                Body = new GetReportRequestBody
                {
                    filialCode = FilialCode,
                    yymm = yymm,
                    reportType = ReportType.T7OldPolis
                }
            };

            var response = Client.GetReport(request)?.Body?.GetReportResult;
            if (response != null)
                return response as ReportT7OldPolis;

            // Получаем константы
            var constantsRequest = new GetT7OldPolisConstantsRequest(
                new GetT7OldPolisConstantsRequestBody(yymm, FilialCode)
            );
            var constantsResponse = Client.GetT7OldPolisConstants(constantsRequest);
            var constants = constantsResponse?.Body?.GetT7OldPolisConstantsResult
                            ?? new ReportT7OldPolisDataDto();

            // Формируем массив данных для всех тем
            var reportList = new List<ReportT7OldPolisDto>();

            foreach (var dictItem in ThemesList) // ← используем reportsDictionary (он же ThemesList)
            {
                var theme = dictItem.Key; // предполагается, что KmsReportDictionary имеет свойство Key
                reportList.Add(new ReportT7OldPolisDto
                {
                    Theme = theme,
                    Data = new ReportT7OldPolisDataDto
                    {
                        Constant2019Count = constants.Constant2019Count,
                        AnnualCount = constants.AnnualCount,
                        CurrentQuantity = 0,
                        CountOldPolis = 0
                    }
                });
            }

            return new ReportT7OldPolis
            {
                Yymm = yymm,
                ReportDataList = reportList.ToArray() // ← массив, как требуется моделью
            };
        }

        public override void FillDataGridView(string form)
        {

            var reportT7OldPolis = Report.ReportDataList.FirstOrDefault(x => x.Theme == form);
            if (reportT7OldPolis == null)
            {
                return;
            }

            if (reportT7OldPolis.Data != null)
            {
                var data = reportT7OldPolis.Data;

                // Столбец 0: постоянное значение (2019)
                Dgv.Rows[0].Cells[0].Value = data.Constant2019Count;

                // Столбец 1: годовой план
                Dgv.Rows[0].Cells[1].Value = data.AnnualCount;

                // Столбец 2: численность
                Dgv.Rows[0].Cells[2].Value = data.CurrentQuantity;

                // Столбец 3: кол-во полисов
                Dgv.Rows[0].Cells[3].Value = data.CountOldPolis;

                // Столбец 4: доля (%) — защита от деления на 0
                if (data.CurrentQuantity != 0)
                {
                    var percentage = (double)data.CountOldPolis / data.CurrentQuantity * 100;
                    Dgv.Rows[0].Cells[4].Value = Math.Round(percentage, 2); // 2 знака после запятой
                }
                else
                {
                    Dgv.Rows[0].Cells[4].Value = 0; // или DBNull.Value, или "—"
                }

                // Столбец 5: динамика за год = факт - годовой план (может быть отрицательной)
                Dgv.Rows[0].Cells[5].Value = data.CountOldPolis - data.AnnualCount;

                // Столбец 6: динамика с 2019 = факт - константа 2019 (может быть отрицательной)
                Dgv.Rows[0].Cells[6].Value = data.CountOldPolis - data.Constant2019Count;
            }

            SetFormula();

            foreach (DataGridViewColumn col in Dgv.Columns)
            {
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }
        }

        public override void SaveReportDataSourceExcel()
        { }
        public override void SaveReportDataSourceHandle()
        { }
        public void SetFormula()
        {
            if (Dgv.Rows.Count == 0 || Dgv.Rows[0].IsNewRow)
                return;

            var row = Dgv.Rows[0];

            // Получаем значения из редактируемых столбцов
            var currentQuantity = GlobalUtils.TryParseInt(row.Cells[2].Value);
            var countOldPolis = GlobalUtils.TryParseInt(row.Cells[3].Value);

            // Получаем константы из недоступных для редактирования столбцов
            var constant2019 = GlobalUtils.TryParseInt(row.Cells[0].Value);
            var annualCount = GlobalUtils.TryParseInt(row.Cells[1].Value);

            // Столбец 4: доля (%)
            if (currentQuantity != 0)
            {
                var percentage = (double)countOldPolis / currentQuantity * 100;
                row.Cells[4].Value = Math.Round(percentage, 2);
            }
            else
            {
                row.Cells[4].Value = 0;
            }

            // Столбец 5: динамика за год
            row.Cells[5].Value = countOldPolis - annualCount;

            // Столбец 6: динамика с 2019
            row.Cells[6].Value = countOldPolis - constant2019;
        }

        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource)
        {

        }

        public override void InitReport()
        {

        }

        public override bool IsVisibleBtnDownloadExcel() => false;

        public override bool IsVisibleBtnHandle() => false;

        public override bool IsVisibleBtnSummary() => false;

        public override void MapForAutoFill(AbstractReport report)
        {
            if (report == null)
            {
                return;
            }
            var inReport = report as ReportT7OldPolis;

            var index = Report.ReportDataList.ToList().FindIndex(x => x.Theme == Cmb.Text);
            var inTheme = inReport.ReportDataList.Single(x => x.Theme == Cmb.Text);
            Report.ReportDataList[index] = inTheme;
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
                    reportType = ReportType.T7OldPolis
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportT7OldPolis;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
        }

        public override void ToExcel(string filename, string filialName)
        {
            //var excel = new ExcelT5NewbornCreator(filename, ExcelForm.T5Newborn, Report.Yymm, filialName, Client, FilialCode);
            //excel.CreateReport(Report, null);
        }

        public override string ValidReport() { return ""; }

        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            Dgv.AllowUserToAddRows = false;
            Dgv.ColumnHeadersVisible = true;

            Dgv.Columns.Clear();
            Dgv.Rows.Clear();

            // Индексы столбцов, которые ТОЛЬКО ДЛЯ ЧТЕНИЯ
            var readOnlyColumns = new HashSet<int> { 0, 1, 4, 5, 6 };

            for (int i = 0; i < based.Count; i++)
            {
                var isReadOnly = readOnlyColumns.Contains(i);
                var backColor = isReadOnly ? Color.LightGray : Color.Azure;

                var column = new DataGridViewTextBoxColumn
                {
                    HeaderText = based[i],
                    Name = $"Col{i}",
                    ReadOnly = isReadOnly, // ← ключевая настройка!
                    SortMode = DataGridViewColumnSortMode.NotSortable,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        BackColor = backColor
                    }
                };

                if (i == 4) // столбец "Доля (%)"
                {
                    column.DefaultCellStyle.Format = "0.00'%'"; 
                }

                Dgv.Columns.Add(column);
            }

            Dgv.Rows.Add();
            Dgv.AutoSize = true;
            Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

        }
        protected override void FillReport(string form)
        {
            var reportRow = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            var row = Dgv.Rows[0];

            // Только столбцы 2 и 3 могут быть изменены пользователем
            reportRow.Data = new ReportT7OldPolisDataDto
            {
                // Защита: не читаем значения из столбцов 0,1,4,5,6
                Constant2019Count = reportRow.Data?.Constant2019Count ?? 0,
                AnnualCount = reportRow.Data?.AnnualCount ?? 0,

                // Только эти поля берутся из UI
                CurrentQuantity = GlobalUtils.TryParseInt(row.Cells[2].Value),
                CountOldPolis = GlobalUtils.TryParseInt(row.Cells[3].Value)
            };

            SetFormula();
        }
    }
}
