using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using NLog.Fluent;

namespace KmsReportClient.Report.Basic
{
    public class ReportIizlProcessor2022 : AbstractReportProcessor<ReportIizl>
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        public ReportIizlProcessor2022(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv,
           ComboBox cmb, TextBox txtb, TabPage page) :
           base(inClient, dgv, cmb, txtb, page,
               XmlFormTemplate.Iizl2022.GetDescription(),
               Log,
               ReportGlobalConst.ReportIizl2022,
               reportsDictionary)
        {
            InitReport();
        }

        public override AbstractReport CollectReportFromWs(string yymm)
        {
            var request = new GetReportRequest
            {
                Body = new GetReportRequestBody { filialCode = FilialCode, yymm = yymm, reportType = ReportType.Iizl2022 }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as ReportIizl;
        }
        public override void FillDataGridView(string form)
        {
            if (form == null)
            {
                return;
            }

            if (form.EndsWith("-Э") || form.EndsWith("-П"))
            {
                FillDgv(Dgv, form);
            }
            else //Согласие
            {
                var reportIizlDto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form) ?? new ReportIizlDto();
                var data = reportIizlDto.Data;

                if (data.Length <= 0)
                {
                    return;
                }

                foreach (DataGridViewRow row in Dgv.Rows)
                {
                    string code = row.Cells[1].Value.ToString();
                    var countPersFirst = data.Single(x => x.Code == code).CountPersFirst;
                    row.Cells[2].Value = countPersFirst;
                }
            }
        }
        public override void SaveReportDataSourceExcel()
        { }
        public override void SaveReportDataSourceHandle()
        { }
        private void FillDgv(DataGridView dgvReport, string form)
        {
            var reportIizlDto = Report.ReportDataList.SingleOrDefault(x => x.Theme.ToLower() == form.ToLower());

            if (reportIizlDto == null)
            {
                return;
            }

            if (reportIizlDto.Data == null || reportIizlDto.Data.Length == 0)
            {
                return;
            }

            if (reportIizlDto != null)
            {

                var reportData = new List<ReportIizlDataDto>();
                foreach (DataGridViewRow row in dgvReport.Rows)
                {
                    var code = row.Cells[1].Value.ToString();

                    var data = reportIizlDto.Data.FirstOrDefault(x => x.Code == code);
                    if (data == null)
                    {
                        continue;
                    }

                    if (form.EndsWith("-Э"))
                    {

                        PgDgvUtils.SetRowText(data.CountPersFirst, row.Cells[2]);
                        PgDgvUtils.SetRowText(data.CountPersRepeat, row.Cells[3]);
                        PgDgvUtils.SetRowText(data.CountMessages, row.Cells[6]);
                        PgDgvUtils.SetRowText(data.TotalCost, row.Cells[7]);
                        PgDgvUtils.SetRowText(data.AverageCostPerMessage, row.Cells[8]);
                        PgDgvUtils.SetRowText(data.AverageCostOfInforming1PL, row.Cells[9]);
                        PgDgvUtils.SetRowText(data.AccountingDocument, row.Cells[10]);

                    }
                    else if (form.EndsWith("-П"))
                    {

                        PgDgvUtils.SetRowText(data.CountPersFirst, row.Cells[2]);
                        PgDgvUtils.SetRowText(data.CountPersRepeat, row.Cells[3]);
                        PgDgvUtils.SetRowText(data.TotalCost, row.Cells[5]);
                        PgDgvUtils.SetRowText(data.AccountingDocument, row.Cells[7]);
                    }


                    reportData.Add(data);


                }

                reportIizlDto.Data = reportData.ToArray();

                SetCalculateCellsValue();
            }
        }


        private void FillReportThemes(DataGridView dgvReport, string form)
        {
            var reportIizlDto = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            if (reportIizlDto != null)
            {

                var reportData = new List<ReportIizlDataDto>();
                foreach (DataGridViewRow row in dgvReport.Rows)
                {
                    var code = row.Cells[1].Value.ToString();

                    //Не сохраняем
                    if (row.Cells[0].Value.ToString().ToLower() == "сумма")
                    {
                        continue;
                    }

                    ReportIizlDataDto data = null;
                    if (form.EndsWith("-Э"))
                    {
                        data = new ReportIizlDataDto
                        {
                            Code = code,
                            CountPersFirst = GlobalUtils.TryParseInt(row.Cells[2].Value),
                            CountPersRepeat = GlobalUtils.TryParseInt(row.Cells[3].Value),
                            CountMessages = GlobalUtils.TryParseInt(row.Cells[6].Value),
                            TotalCost = GlobalUtils.TryParseDecimal(row.Cells[7].Value),
                            AverageCostPerMessage = GlobalUtils.TryParseDecimal(row.Cells[8].Value),
                            AverageCostOfInforming1PL = GlobalUtils.TryParseDecimal(row.Cells[9].Value),
                            AccountingDocument = row.Cells[10].Value?.ToString() ?? ""
                        };
                    }
                    else if (form.EndsWith("-П"))
                    {
                        data = new ReportIizlDataDto
                        {
                            Code = code,
                            CountPersFirst = GlobalUtils.TryParseInt(row.Cells[2].Value),
                            CountPersRepeat = GlobalUtils.TryParseInt(row.Cells[3].Value),
                            TotalCost = GlobalUtils.TryParseDecimal(row.Cells[5].Value),
                            AccountingDocument = row.Cells[7].Value?.ToString() ?? ""
                        };
                    }


                    reportData.Add(data);


                }

                reportIizlDto.Data = reportData.ToArray();
            }
        }

        private void FillInfAgreeReport(DataGridView dgvReport, string form)
        {
            var reportIizlDto = Report.ReportDataList.Single(x => x.Theme == form);
            if (reportIizlDto == null)
            {
                return;
            }
            reportIizlDto.Data = (from DataGridViewRow row in dgvReport.Rows
                                  let code = row.Cells[1].Value.ToString()
                                  let sum = row.Cells[2].Value?.ToString() ?? "0"
                                  select new ReportIizlDataDto { Code = code, CountPersFirst = GlobalUtils.TryParseInt(sum) }).ToArray();
        }


        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource)
        {

        }
        public override void InitReport()
        {
            Report = new ReportIizl { ReportDataList = new ReportIizlDto[ThemesList.Count], IdType = IdReportType };

            int i = 0;
            foreach (var theme in ThemesList.Select(x => x.Key))
            {
                var themeData = ThemeTextData.Tables_fromxml.Single(x => x.TableName_fromxml == theme);
                var rows = themeData.Rows_fromxml.Select(x => new ReportIizlDataDto { Code = x.RowNum_fromxml }).ToArray();

                Report.ReportDataList[i++] = new ReportIizlDto { Theme = theme, Data = rows };
            }
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

            var inReport = report as ReportIizl;

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
                    reportType = ReportType.Iizl2022
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportIizl;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
        }
        public override void ToExcel(string filename, string filialName)
        {
            var mm = YymmUtils.GetMonth(Report.Yymm.Substring(2, 2)) + " 20" + Report.Yymm.Substring(0, 2);
            ExcelIizl2022Creator excel = new ExcelIizl2022Creator(filename, ExcelForm.iizl2022, mm, filialName);
            excel.CreateReport(Report, null);
        }
        public override string ValidReport()
        {
            return "";
        }
        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            if (form.EndsWith("-Э"))
            {
                CreateDgvThemeColumnsElectronicMeans(Dgv);
            }
            else if (form.EndsWith("-П"))
            {
                CreateDgvThemeColumnsWrittenInformation(Dgv);
            }
            else
            {
                CreateDgvInfAgreeColumns(Dgv);
            }

            int countRows = ThemeTextData.Tables_fromxml.Single(x => x.TableName_fromxml == form).RowsCount_fromxml;

            foreach (var row in table)
            {
                var dgvRow = new DataGridViewRow();
                var cellName = new DataGridViewTextBoxCell { Value = row.RowText_fromxml };
                var cellNum = new DataGridViewTextBoxCell { Value = row.RowNum_fromxml };
                dgvRow.Cells.Add(cellName);
                dgvRow.Cells.Add(cellNum);

                if (row.RowText_fromxml == "Сумма")
                {
                    dgvRow.ReadOnly = true;
                    dgvRow.DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.LightGray };
                }

                var exclusionCells = row.ExclusionCells_fromxml?.Split(',');

                for (int i = 2; i < countRows; i++)
                {
                    bool isNeedExcludeSum = exclusionCells?.Contains(i.ToString()) ?? false;
                    var cell = new DataGridViewTextBoxCell { Value = isNeedExcludeSum ? "x" : "0" };
                    dgvRow.Cells.Add(cell);

                    if (isNeedExcludeSum)
                    {
                        cell.ReadOnly = true;
                        cell.Style.BackColor = Color.DarkGray;
                    }
                }

                Dgv.Rows.Add(dgvRow);
            }
        }

        private void CreateDgvInfAgreeColumns(DataGridView dgvReport)
        {
            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Наименование",
                Width = 450,
                DataPropertyName = "Naim",
                Name = "Naim",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Номер строки",
                Width = 100,
                DataPropertyName = "Num",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Name = "Num",
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Количество",
                Width = 100,
                DataPropertyName = "Code",
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Name = "Code"
            };
            dgvReport.Columns.Add(column);
        }

        private void CreateDgvThemeColumnsWrittenInformation(DataGridView dgvReport)
        {
            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Способы информирования",
                Width = 265,
                DataPropertyName = "Way",
                Name = "Way",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Код",
                Width = 50,
                DataPropertyName = "Code",
                Name = "Code",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Количество ЗЛ (первичное информирование по теме)" + Environment.NewLine + @"(гр.1)",
                Width = 120,
                DataPropertyName = "CountPeopleFirst",
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Name = "CountPeopleFirst"
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Количество ЗЛ (повторное информирование по теме)" + Environment.NewLine + @"(гр.2)",
                Width = 120,
                DataPropertyName = "CountPeopleRepeat",
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Name = "CountPeopleRepeat"
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Количество сообщений (первичное и повторное информирование по теме)" +
                             Environment.NewLine + @"(гр.3)" +
                             Environment.NewLine + @"гр.1 + гр.2"
,
                Width = 120,
                DataPropertyName = "CountMessages",
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Name = "CountMessages",
                ReadOnly = true,
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.LightGray }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Суммарные затраты(руб.)" + Environment.NewLine + @"(гр.4)",
                Width = 120,
                DataPropertyName = "TotalCost",
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Name = "TotalCost"
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Средняя стоимость сообщения (руб.)Средняя стоимость сообщения (руб.)" +
                             Environment.NewLine + @"(гр.5)" +
                             Environment.NewLine + @"гр.4 / гр.3",
                Width = 120,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                ReadOnly = true,
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.LightGray }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Реквизиты учетного документа" + Environment.NewLine + @"(гр.5)",
                Width = 120,
                DataPropertyName = "AccountingDocument",
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Name = "AccountingDocument"
            };
            dgvReport.Columns.Add(column);
        }

        private void CreateDgvThemeColumnsElectronicMeans(DataGridView dgvReport)
        {
            var column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Способы информирования",
                Width = 265,
                DataPropertyName = "Way",
                Name = "Way",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Код",
                Width = 50,
                DataPropertyName = "Code",
                Name = "Code",
                ReadOnly = true,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Azure }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Количество ЗЛ (первичное информирование по теме)" + Environment.NewLine + @"(гр.1)",
                Width = 120,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Количество ЗЛ (повторное информирование по теме)" + Environment.NewLine + @"(гр.2)",
                Width = 120,
                DataPropertyName = "CountPeopleRepeat",
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Итого сообщений (1 сообщение = 1 ЗЛ)(первичное и повторное информирование по теме)" +
                             Environment.NewLine + @"(гр.3)" +
                             Environment.NewLine + @"гр.1 + гр.2",
                Width = 120,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                ReadOnly = true,
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.LightGray }

            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Среднее количество сегментов на 1  ЗЛ" +
                            Environment.NewLine + @"(гр.4)" +
                            Environment.NewLine + @"гр.5 / гр.3",
                Width = 120,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                ReadOnly = true,
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.LightGray }
            };
            dgvReport.Columns.Add(column);
            column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Количество сегментов сообщений (общее)" +
                           Environment.NewLine + @"(гр.5)",
                Width = 120,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            dgvReport.Columns.Add(column);

            column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Суммарные затраты(руб.)" + Environment.NewLine + @"(гр.6)",
                Width = 120,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            dgvReport.Columns.Add(column);

            column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Средняя стоимость сегмента сообщения (руб.)" +
                Environment.NewLine + @"(гр.7)" +
                Environment.NewLine + @"гр.6 / гр.5",
                Width = 120,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                ReadOnly = !(GetCurrentTheme().Contains("Тема С-Э") || GetCurrentTheme().Contains("Тема К-Э")),
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = !(GetCurrentTheme().Contains("Тема С-Э") || GetCurrentTheme().Contains("Тема К-Э")) ? Color.LightGray : Color.White }
            };
            dgvReport.Columns.Add(column);

            column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Средняя стоимость информирования 1 ЗЛ" +
               Environment.NewLine + @"(гр.8)" +
               Environment.NewLine + @"гр.4 * гр.7",
                Width = 120,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                ReadOnly = !(GetCurrentTheme().Contains("Тема С-Э") || GetCurrentTheme().Contains("Тема К-Э")),
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = !(GetCurrentTheme().Contains("Тема С-Э") || GetCurrentTheme().Contains("Тема К-Э")) ? Color.LightGray : Color.White }
            };
            dgvReport.Columns.Add(column);

            column = new DataGridViewTextBoxColumn
            {
                HeaderText = @"Реквизиты учетного документа - файл со списком рассылки (по теме иные в примечании указать тему)" + Environment.NewLine + @"(гр.9)",
                Width = 120,
                SortMode = DataGridViewColumnSortMode.NotSortable,
            };
            dgvReport.Columns.Add(column);
        }

        protected override void FillReport(string form)
        {
            if (form == null)
            {
                return;
            }

            if (form.EndsWith("-Э") || form.EndsWith("-П"))
            {
                FillReportThemes(Dgv, form);
            }
            else
            {
                FillInfAgreeReport(Dgv, form);
            }
        }

        public void SetCalculateCellsValue()
        {
            try
            {
                //Итого
                string theme = GetCurrentTheme();
                if (theme.Contains("Согл"))
                    return;




                foreach (DataGridViewRow row in Dgv.Rows)
                {
                    if (row.Cells[1].Value.ToString().StartsWith("С"))
                    {
                        continue;
                    }
                    if (theme.EndsWith("-Э"))
                    {
                        try
                        {
                            //Электронные средства столбец 3 ( 1+2 )
                            PgDgvUtils.SetRowText(GlobalUtils.TryParseDecimal(row.Cells[2].Value) + GlobalUtils.TryParseDecimal(row.Cells[3].Value), row.Cells[4]);

                        }
                        catch { }

                        //Электронные средства столбец 4 ( 5/3 )
                        try
                        {
                            if (GlobalUtils.TryParseDecimal(row.Cells[4].Value) != 0)
                            {
                                PgDgvUtils.SetRowText(Math.Round(GlobalUtils.TryParseDecimal(row.Cells[6].Value) / GlobalUtils.TryParseDecimal(row.Cells[4].Value), 2), row.Cells[5]);

                            }

                        }
                        catch { }


                        if (!(theme.Contains("Тема С-Э") || theme.Contains("Тема К-Э")))
                        {
                            //Электронные средства столбец 7 ( 6/5 )
                            try
                            {
                                if (GlobalUtils.TryParseDecimal(row.Cells[6].Value) != 0)
                                {
                                    PgDgvUtils.SetRowText(Math.Round(GlobalUtils.TryParseDecimal(row.Cells[7].Value) / GlobalUtils.TryParseDecimal(row.Cells[6].Value), 2), row.Cells[8]);

                                }


                            }
                            catch { }

                            //Электронные средства столбец 8 ( 4*7 )
                            try
                            {
                                if (GlobalUtils.TryParseDecimal(row.Cells[8].Value) != 0)
                                {

                                    PgDgvUtils.SetRowText(Math.Round(GlobalUtils.TryParseDecimal(row.Cells[5].Value) * GlobalUtils.TryParseDecimal(row.Cells[8].Value), 2), row.Cells[9]);
                                }

                            }
                            catch { }
                        }



                    }

                    if (theme.EndsWith("-П"))
                    {
                        try
                        {
                            PgDgvUtils.SetRowText(GlobalUtils.TryParseDecimal(row.Cells[2].Value) + GlobalUtils.TryParseDecimal(row.Cells[3].Value), row.Cells[4]);

                        }
                        catch { }

                        try
                        {
                            if (GlobalUtils.TryParseDecimal(row.Cells[4].Value) != 0)
                                PgDgvUtils.SetRowText(Math.Round(GlobalUtils.TryParseDecimal(row.Cells[5].Value) / GlobalUtils.TryParseDecimal(row.Cells[4].Value), 2), row.Cells[6]);
                        }
                        catch { }


                    }

                }


                //Сумма
                var totalRow = Dgv.Rows[Dgv.Rows.Count - 1];
                foreach (DataGridViewCell totalCell in totalRow.Cells)
                {
                    if (totalCell.Value.ToString() == "x")
                        continue;

                    if (totalCell.ColumnIndex == 0 || totalCell.ColumnIndex == 1)
                        continue;

                    totalCell.Value = Dgv.Rows.OfType<DataGridViewRow>().Where(x => x.Index != totalRow.Index).Sum(x => GlobalUtils.TryParseDecimal(x.Cells[totalCell.ColumnIndex].Value));
                }



            }
            catch (Exception ex)
            {

            }
        }
    }
}
