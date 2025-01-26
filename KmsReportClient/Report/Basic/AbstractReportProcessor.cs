using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Xml.Serialization;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Model.XML;
using KmsReportClient.Support;
using NLog;

namespace KmsReportClient.Report.Basic
{
    public abstract class AbstractReportProcessor<TR> : IReportProcessor where TR : AbstractReport
    {
        protected readonly EndpointSoap Client;

        protected readonly ComboBox Cmb;
        protected readonly DataGridView Dgv;
        protected readonly TextBox Txtb;
        protected readonly TabPage Page;

        protected readonly string IdReportType;
        protected readonly string SerializeName;
        protected readonly string SmallName;
        protected readonly TemplateForm ThemeTextData;
        protected readonly Logger log;

        protected string FilialCode;
        protected string OldTheme;
        protected string FilialName;
        protected bool HasReport;
        protected Color ColorReport;
        protected TR Report;

        protected AbstractReportProcessor(
            EndpointSoap client,
            DataGridView dgv,
            ComboBox cmb,
            TextBox txtb,
            TabPage page,
            string filename,
            Logger logger,
            string id,
            List<KmsReportDictionary> reportsDictionary
            )
        {
            Client = client;
            Dgv = dgv;
            Txtb = txtb;
            Page = page;
            log = logger;
            IdReportType = id;
            Cmb = cmb;

            SmallName = reportsDictionary.Single(x => x.Key == id).Value;
            SerializeName = id + "Data";
            ThemeTextData = ReadTemplateXml(AppDomain.CurrentDomain.BaseDirectory + "Template\\" + filename);

            var themes = GetForms();
            Cmb.DataSource = themes;
            Cmb.DisplayMember = "Key";
            Cmb.ValueMember = "Value";

            txtb.Text = themes[0].Value;
        }

        public abstract AbstractReport CollectReportFromWs(string yymm);
        public abstract void FillDataGridView(string form);
        public abstract void SaveToDb();
        public abstract void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource);
        public abstract void ToExcel(string filename, string filialName);
        public abstract string ValidReport();
        public abstract bool IsVisibleBtnDownloadExcel();
        public abstract bool IsVisibleBtnHandle();
        public abstract bool IsVisibleBtnSummary();
        public abstract void InitReport();
        public abstract void MapForAutoFill(AbstractReport report);
        public abstract void SaveReportDataSourceHandle();
        public abstract void SaveReportDataSourceExcel();


        public virtual void CalculateCells()
        {
            Console.WriteLine();
        }


        public void CreateTotalColumn()
        {
            if (GetCurrentTheme() != "Результаты МЭК" && GetCurrentTheme() != "Таблица 5А" && GetCurrentTheme() != "Оплата МП")
            {
                if (Report.IdType == "Zpz10" || Report.IdType == "Zpz10_2025") { Dgv.Columns.Add("Total", "С начала года"); }
                else { Dgv.Columns.Add("Total", "Итого"); }

                Dgv.Columns["Total"].ReadOnly = true;
                Dgv.Columns["Total"].DefaultCellStyle.BackColor = Color.LightGray;

                Dgv.Columns["Total"].DisplayIndex = 2;

                if (Report.IdType == "PG" || Report.IdType == "PG_Q")
                {
                    if (GetCurrentTheme() == "Таблица 6" || GetCurrentTheme() == "Таблица 8")
                    {
                        Dgv.Columns["Total"].HeaderText = "Итого цел";

                        Dgv.Columns.Add("TotalPlan", "Итого план");
                        Dgv.Columns["TotalPlan"].DisplayIndex = 9;

                        Dgv.Columns["TotalPlan"].ReadOnly = true;
                        Dgv.Columns["TotalPlan"].DefaultCellStyle.BackColor = Color.LightGray;

                        Dgv.Columns.Add("TotalPlanCel", "Итого");
                        Dgv.Columns["TotalPlanCel"].DisplayIndex = 2;

                        Dgv.Columns["TotalPlanCel"].ReadOnly = true;
                        Dgv.Columns["TotalPlanCel"].DefaultCellStyle.BackColor = Color.Gray;
                    }
                }
                if (Report.IdType == "Zpz" || Report.IdType == "Zpz_Q" || Report.IdType == "Zpz2025" || Report.IdType == "Zpz_Q2025")
                {
                    if (GetCurrentTheme() == "Таблица 6" || GetCurrentTheme() == "Таблица 7")
                    {
                        Dgv.Columns["Total"].HeaderText = "Итого цел";

                        Dgv.Columns.Add("TotalPlan", "Итого план");
                        Dgv.Columns["TotalPlan"].DisplayIndex = 9;

                        Dgv.Columns["TotalPlan"].ReadOnly = true;
                        Dgv.Columns["TotalPlan"].DefaultCellStyle.BackColor = Color.LightGray;
                        Dgv.Columns["Total"].DefaultCellStyle.BackColor = Color.LightGray;
                        Dgv.Columns.Add("TotalPlanCel", "Итого");
                        Dgv.Columns["TotalPlanCel"].DisplayIndex = 2;

                        Dgv.Columns["TotalPlanCel"].ReadOnly = true;
                        Dgv.Columns["TotalPlanCel"].DefaultCellStyle.BackColor = Color.LightGreen;
                    }
                    if (GetCurrentTheme() == "Таблица 8" && Report.IdType == "Zpz_Q2025")
                    {
                        Dgv.Columns["Total"].HeaderText = "Итого результаты";
                        Dgv.Columns["Total"].DefaultCellStyle.BackColor = Color.LightGray;
                        Dgv.Columns["Total"].ReadOnly = true;
                    }
                }
            }
        }

        public void SetTotalColumn()
        {
            try
            {
                int columnCount = Dgv.Columns.Count - 1;
                for (int row = 0; row < Dgv.Rows.Count; row++)
                {
                    Dgv.Rows[row].Cells[columnCount].Value = 0;
                    int valueCel = 0;
                    int valuePlan = 0;
                    decimal valueSMO = 0;

                    for (int cell = 1; cell < Dgv.Rows[row].Cells.Count - 1; cell++)
                    {
                        if (Dgv.Rows[row].Cells[cell].Value == null)
                            continue;

                        if (!Dgv.Columns[Dgv.Rows[row].Cells[cell].ColumnIndex].Name.Contains("Row") && Dgv.Rows[row].Cells[cell].Value.ToString() != "x")
                        {
                            //Console.WriteLine($"{GetCurrentTheme()} {Report.IdType}");
                            if ((Report.IdType == "PG" || Report.IdType == "PG_Q") && (GetCurrentTheme() == "Таблица 5" || GetCurrentTheme() == "Таблица 6" || GetCurrentTheme() == "Таблица 8"))
                            {
                                if (Dgv.Rows[row].Cells[cell].ColumnIndex == 2 || Dgv.Rows[row].Cells[cell].ColumnIndex == 3 || Dgv.Rows[row].Cells[cell].ColumnIndex == 4 || Dgv.Rows[row].Cells[cell].ColumnIndex == 6)
                                {
                                    valueCel += GlobalUtils.TryParseInt(Dgv.Rows[row].Cells[cell].Value);
                                }
                                else if (Dgv.Rows[row].Cells[cell].ColumnIndex == 8 || Dgv.Rows[row].Cells[cell].ColumnIndex == 9 || Dgv.Rows[row].Cells[cell].ColumnIndex == 10 || Dgv.Rows[row].Cells[cell].ColumnIndex == 12)
                                {
                                    valuePlan += GlobalUtils.TryParseInt(Dgv.Rows[row].Cells[cell].Value);
                                }
                                /// НА ДАННОМ ЭТАПЕ ПРОБЕГАЕМ ДЛЯ КАЖДОЙ СТРОКИ ОТЧЕТА ПО ЯЧЕЙКАМ и СУММИРУЕМ ЗНАЧЕНИЯ ДЛЯ ЦЕЛЕВЫХ (2-3-4-6) И ПЛАНОВЫХ (8-9-10-12), ПОКА ПРОСТО В ПЕРЕМЕННЫЕ valueCel и valuePlan ///
                            }
                            else if ((Report.IdType == "Zpz" || Report.IdType == "Zpz_Q" || Report.IdType == "Zpz2025" || Report.IdType == "Zpz_Q2025") && (GetCurrentTheme() == "Таблица 5А" || GetCurrentTheme() == "Результаты МЭК" || GetCurrentTheme() == "Таблица 6" || GetCurrentTheme() == "Таблица 7"))
                            {
                                if (Dgv.Rows[row].Cells[cell].ColumnIndex == 2 || Dgv.Rows[row].Cells[cell].ColumnIndex == 3 || Dgv.Rows[row].Cells[cell].ColumnIndex == 4 || Dgv.Rows[row].Cells[cell].ColumnIndex == 6)
                                {
                                    valueCel += GlobalUtils.TryParseInt(Dgv.Rows[row].Cells[cell].Value);
                                }
                                else if (Dgv.Rows[row].Cells[cell].ColumnIndex == 8 || Dgv.Rows[row].Cells[cell].ColumnIndex == 9 || Dgv.Rows[row].Cells[cell].ColumnIndex == 10 || Dgv.Rows[row].Cells[cell].ColumnIndex == 12)
                                {
                                    valuePlan += GlobalUtils.TryParseInt(Dgv.Rows[row].Cells[cell].Value);
                                }
                                /// НА ДАННОМ ЭТАПЕ ПРОБЕГАЕМ ДЛЯ КАЖДОЙ СТРОКИ ОТЧЕТА ПО ЯЧЕЙКАМ и СУММИРУЕМ ЗНАЧЕНИЯ ДЛЯ ЦЕЛЕВЫХ (2-3-4-6) И ПЛАНОВЫХ (8-9-10-12), ПОКА ПРОСТО В ПЕРЕМЕННЫЕ valueCel и valuePlan ///
                            }

                            else if ((Report.IdType == "Zpz_Q2025") && (GetCurrentTheme() == "Таблица 8"))
                            {
                                if (Dgv.Rows[row].Cells[cell].ColumnIndex == 2 || Dgv.Rows[row].Cells[cell].ColumnIndex == 3 || Dgv.Rows[row].Cells[cell].ColumnIndex == 4 || Dgv.Rows[row].Cells[cell].ColumnIndex == 6)
                                {
                                    valueSMO += GlobalUtils.TryParseDecimal(Dgv.Rows[row].Cells[cell].Value);
                                }
                                /// НА ДАННОМ ЭТАПЕ ПРОБЕГАЕМ ДЛЯ КАЖДОЙ СТРОКИ ОТЧЕТА ПО ЯЧЕЙКАМ и СУММИРУЕМ ЗНАЧЕНИЯ (2-3-4-6) ПОКА ПРОСТО В ПЕРЕМЕННУЮ valueSMO ///
                            }

                            else if ((Report.IdType == "Zpz_Q2025") && (GetCurrentTheme() == "Таблица 9"))
                            {
                                if (Dgv.Rows[row].Cells[cell].ColumnIndex == 2 || Dgv.Rows[row].Cells[cell].ColumnIndex == 3)
                                {
                                    valueSMO += GlobalUtils.TryParseDecimal(Dgv.Rows[row].Cells[cell].Value);
                                }
                                /// НА ДАННОМ ЭТАПЕ ПРОБЕГАЕМ ДЛЯ КАЖДОЙ СТРОКИ ОТЧЕТА ПО ЯЧЕЙКАМ и СУММИРУЕМ ЗНАЧЕНИЯ (2-3-4-6) ПОКА ПРОСТО В ПЕРЕМЕННУЮ valueSMO ///
                            }

                            else
                            {
                                if ((GetCurrentTheme() == "Таблица 1Л" || GetCurrentTheme() == "Таблица 2Л"))
                                {
                                    if (Dgv.Rows[row].Cells[cell].ColumnIndex != 6)
                                    {
                                        valueCel += GlobalUtils.TryParseInt(Dgv.Rows[row].Cells[cell].Value);

                                    }
                                    // Потребовали, чтобы в 6й колонке была сумма 2-5. Попробуем плюсануть тут.
                                    else if (GetCurrentTheme() == "Таблица 1Л" && row != Dgv.Rows.Count - 1 && row != Dgv.Rows.Count - 2)
                                    { Dgv.Rows[row].Cells[cell].Value = valueCel; };
                                }
                                else if (GetCurrentTheme() == "Таблица 12")
                                {
                                    if (Dgv.Rows[row].Cells[cell].ColumnIndex != 3)
                                    {
                                        valueCel += GlobalUtils.TryParseInt(Dgv.Rows[row].Cells[cell].Value);
                                    }
                                }
                                else
                                {
                                    if ((Report.IdType == "Zpz" || Report.IdType == "Zpz2025") && (GetCurrentTheme() == "Таблица 1" || GetCurrentTheme() == "Таблица 2" || GetCurrentTheme() == "Таблица 3" || GetCurrentTheme() == "Таблица 4"))
                                    {
                                        if (Dgv.Rows[row].Cells[cell].ColumnIndex != 4) { valueCel += GlobalUtils.TryParseInt(Dgv.Rows[row].Cells[cell].Value); }
                                    }
                                    else { valueCel += GlobalUtils.TryParseInt(Dgv.Rows[row].Cells[cell].Value); }
                                }
                            }
                        }
                    }

                    // Разраб до меня сделал так, переделывать не стали, работает и ладно.
                    //// Тот, кто это видит прошу меня простить.
                    if ((Report.IdType == "PG" || Report.IdType == "PG_Q") && (GetCurrentTheme() == "Таблица 6" || GetCurrentTheme() == "Таблица 8"))
                    {
                        Dgv.Rows[row].Cells["Total"].Value = valueCel; //Целевые
                        Dgv.Rows[row].Cells["TotalPlan"].Value = valuePlan; // Плановые
                        Dgv.Rows[row].Cells["TotalPlanCel"].Value = valuePlan + valueCel; // Итого цел + план
                        // Пишем в DGV значения и сумму
                    }
                    else if ((Report.IdType == "Zpz" || Report.IdType == "Zpz_Q" || Report.IdType == "Zpz2025" || Report.IdType == "Zpz_Q2025") && (GetCurrentTheme() == "Таблица 6" || GetCurrentTheme() == "Таблица 7"))
                    {
                        Dgv.Rows[row].Cells["Total"].Value = valueCel; //Целевые
                        Dgv.Rows[row].Cells["TotalPlan"].Value = valuePlan; // Плановые
                        Dgv.Rows[row].Cells["TotalPlanCel"].Value = valuePlan + valueCel; // Итого цел + план
                        // Пишем в DGV значения и сумму
                    }
                    else if ((Report.IdType == "Zpz_Q2025") && (GetCurrentTheme() == "Таблица 8" || GetCurrentTheme() == "Таблица 9"))
                    {
                        Dgv.Rows[row].Cells["Total"].Value = valueSMO;
                        // Пишем в DGV значения и сумму
                    }
                    else
                    {
                        Dgv.Rows[row].Cells[columnCount].Value = valueCel; //Целевые
                    }
                }

                string[] rowFor6Row = { "6.1", "6.2", "6.3", "6.4", "6.5", "6.6", "6.7", "6.8", "6.9", "6.10" };
                if ((Report.IdType == "PG" && (GetCurrentTheme() == "Таблица 6" || GetCurrentTheme() == "Таблица 8" || GetCurrentTheme() == "Таблица 5" || GetCurrentTheme() == "Таблица 10")) ||
                    (Report.IdType == "PG_Q"))
                {
                    var row6 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "6");
                    if (row6 != null)
                    {
                        var rowsForCalcluate6Total = Dgv.Rows.Cast<DataGridViewRow>().Where(x => rowFor6Row.Contains(x.Cells[1].Value.ToString()));

                        row6.Cells["Total"].Value = rowsForCalcluate6Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["Total"].Value));
                        row6.Cells["TotalPlan"].Value = rowsForCalcluate6Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["TotalPlan"].Value));

                        row6.Cells["TotalPlanCel"].Value = GlobalUtils.TryParseDecimal(row6.Cells["Total"].Value) + GlobalUtils.TryParseDecimal(row6.Cells["TotalPlan"].Value); // Итого цел + план
                    }
                    else
                    {
                        var rowsForCalcluate6Total = Dgv.Rows.Cast<DataGridViewRow>().Where(x => "6".Contains(x.Cells[1].Value.ToString()));

                        row6.Cells["Total"].Value = rowsForCalcluate6Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["Total"].Value));
                        row6.Cells["TotalPlan"].Value = rowsForCalcluate6Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["TotalPlan"].Value));

                        row6.Cells["TotalPlanCel"].Value = GlobalUtils.TryParseDecimal(row6.Cells["Total"].Value) + GlobalUtils.TryParseDecimal(row6.Cells["TotalPlan"].Value); // Итого цел + план
                    };
                    /// ПОДСЧЕТ ИТОГОВЫХ СТРОК ДЛЯ ТАБЛИЦЫ 10 - 13 
                    if ((Report.IdType == "PG_Q" || Report.IdType == "PG") && (GetCurrentTheme() == "Таблица 10" || GetCurrentTheme() == "Таблица 13"))
                    {
                        try
                        {
                            string[] rowFor4Row = { "4.1", "4.2", "4.3", "4.4" };
                            var row4 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "4");
                            var rowsForCalcluate4Total = Dgv.Rows.Cast<DataGridViewRow>().Where(x => rowFor4Row.Contains(x.Cells[1].Value.ToString()));
                            row4.Cells["Total"].Value = rowsForCalcluate4Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["Total"].Value));
                            row4.Cells["TotalPlan"].Value = rowsForCalcluate4Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["TotalPlan"].Value));

                            row4.Cells["TotalPlanCel"].Value = GlobalUtils.TryParseDecimal(row4.Cells["Total"].Value) + GlobalUtils.TryParseDecimal(row4.Cells["TotalPlan"].Value); // Итого цел + план
                        }
                        catch (Exception ex)
                        { }

                        try
                        {
                            string[] rowFor5Row = { "5.1", "5.2", "5.3", "5.4", "5.5", "5.6", "5.7", "5.8" };
                            var row5 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "5");
                            var rowsForCalcluate5Total = Dgv.Rows.Cast<DataGridViewRow>().Where(x => rowFor5Row.Contains(x.Cells[1].Value.ToString()));
                            row5.Cells["Total"].Value = rowsForCalcluate5Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["Total"].Value));
                            row5.Cells["TotalPlan"].Value = rowsForCalcluate5Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["TotalPlan"].Value));

                            row5.Cells["TotalPlanCel"].Value = GlobalUtils.TryParseDecimal(row5.Cells["Total"].Value) + GlobalUtils.TryParseDecimal(row5.Cells["TotalPlan"].Value); // Итого цел + план
                        }
                        catch (Exception ex)
                        { }
                        try
                        {
                            string[] rowFor1Row = { "1.1", "1.2", "1.3", "1.4", "1.5" };
                            var row1 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "1");
                            var rowsForCalcluate1Total = Dgv.Rows.Cast<DataGridViewRow>().Where(x => rowFor1Row.Contains(x.Cells[1].Value.ToString()));
                            row1.Cells["Total"].Value = rowsForCalcluate1Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["Total"].Value));
                            row1.Cells["TotalPlan"].Value = rowsForCalcluate1Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["TotalPlan"].Value));

                            row1.Cells["TotalPlanCel"].Value = GlobalUtils.TryParseDecimal(row1.Cells["Total"].Value) + GlobalUtils.TryParseDecimal(row1.Cells["TotalPlan"].Value); // Итого цел + план
                        }
                        catch (Exception ex)
                        { }
                        try
                        {
                            string[] rowFor2Row = { "2.1", "2.2", "2.3", "2.4", "2.5", "2.6", "2.7", "2.8" };
                            var row2 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "2");
                            var rowsForCalcluate2Total = Dgv.Rows.Cast<DataGridViewRow>().Where(x => rowFor2Row.Contains(x.Cells[1].Value.ToString()));
                            row2.Cells["Total"].Value = rowsForCalcluate2Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["Total"].Value));
                            row2.Cells["TotalPlan"].Value = rowsForCalcluate2Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["TotalPlan"].Value));

                            row2.Cells["TotalPlanCel"].Value = GlobalUtils.TryParseDecimal(row2.Cells["Total"].Value) + GlobalUtils.TryParseDecimal(row2.Cells["TotalPlan"].Value); // Итого цел + план
                        }
                        catch (Exception ex)
                        { }
                    }
                    /// ПОДСЧЕТ ИТОГОВЫХ СТРОК ДЛЯ ТАБЛИЦЫ 1 
                    if (Report.IdType == "PG_Q" && GetCurrentTheme() == "Таблица 1")
                    {
                        try
                        {
                            string[] rowFor3Row = { "3.1", "3.2", "3.3", "3.4", "3.5", "3.6", "3.7", "3.8", "3.9", "3.10", "3.11", "3.12" };
                            var row3 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "3");
                            var rowsForCalcluate3Total = Dgv.Rows.Cast<DataGridViewRow>().Where(x => rowFor3Row.Contains(x.Cells[1].Value.ToString()));
                            row3.Cells["Total"].Value = rowsForCalcluate3Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["Total"].Value));
                            row3.Cells["TotalPlan"].Value = rowsForCalcluate3Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["TotalPlan"].Value));

                            row3.Cells["TotalPlanCel"].Value = GlobalUtils.TryParseDecimal(row3.Cells["Total"].Value) + GlobalUtils.TryParseDecimal(row3.Cells["TotalPlan"].Value); // Итого цел + план
                        }
                        catch (Exception ex)
                        { }
                        try
                        {
                            string[] rowFor4Row = { "4.1", "4.2", "4.3", "4.4", "4.5", "4.6", "4.7", "4.8", "4.9", "4.10", "4.11", "4.12", "4.13" };
                            var row4 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "4");
                            var rowsForCalcluate4Total = Dgv.Rows.Cast<DataGridViewRow>().Where(x => rowFor4Row.Contains(x.Cells[1].Value.ToString()));
                            row4.Cells["Total"].Value = rowsForCalcluate4Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["Total"].Value));
                            row4.Cells["TotalPlan"].Value = rowsForCalcluate4Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["TotalPlan"].Value));

                            row4.Cells["TotalPlanCel"].Value = GlobalUtils.TryParseDecimal(row4.Cells["Total"].Value) + GlobalUtils.TryParseDecimal(row4.Cells["TotalPlan"].Value); // Итого цел + план
                        }
                        catch (Exception ex)
                        { }
                        try
                        {
                            string[] rowFor1Row = { "2", "4", "5" };
                            var row1 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "1");
                            var rowsForCalcluate1Total = Dgv.Rows.Cast<DataGridViewRow>().Where(x => rowFor1Row.Contains(x.Cells[1].Value.ToString()));
                            row1.Cells["Total"].Value = rowsForCalcluate1Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["Total"].Value));
                            row1.Cells["TotalPlan"].Value = rowsForCalcluate1Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["TotalPlan"].Value));

                            row1.Cells["TotalPlanCel"].Value = GlobalUtils.TryParseDecimal(row1.Cells["Total"].Value) + GlobalUtils.TryParseDecimal(row1.Cells["TotalPlan"].Value); // Итого цел + план
                        }
                        catch (Exception ex)
                        { }
                    }
                    /// ПОДСЧЕТ ИТОГОВЫХ СТРОК ДЛЯ ТАБЛИЦЫ 3 
                    if (Report.IdType == "PG_Q" && GetCurrentTheme() == "Таблица 3")
                    {
                        try
                        {
                            string[] rowFor1Row = { "1.1", "2.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9", "1.10", "1.11", "1.12" };
                            var row1 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "1");
                            var rowsForCalcluate1Total = Dgv.Rows.Cast<DataGridViewRow>().Where(x => rowFor1Row.Contains(x.Cells[1].Value.ToString()));
                            row1.Cells["Total"].Value = rowsForCalcluate1Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["Total"].Value));
                            row1.Cells["TotalPlan"].Value = rowsForCalcluate1Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["TotalPlan"].Value));

                            row1.Cells["TotalPlanCel"].Value = GlobalUtils.TryParseDecimal(row1.Cells["Total"].Value) + GlobalUtils.TryParseDecimal(row1.Cells["TotalPlan"].Value); // Итого цел + план
                        }
                        catch (Exception ex)
                        { }
                    }
                    /// ПОДСЧЕТ ИТОГОВЫХ СТРОК ДЛЯ ТАБЛИЦЫ 5 
                    if (Report.IdType == "PG_Q" && GetCurrentTheme() == "Таблица 5")
                    {
                        try
                        {
                            string[] rowFor4Row = { "4.1", "4.2", "4.3", "4.4", "4.5", "4.6" };
                            var row4 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "4");
                            var rowsForCalcluate4Total = Dgv.Rows.Cast<DataGridViewRow>().Where(x => rowFor4Row.Contains(x.Cells[1].Value.ToString()));
                            row4.Cells["Total"].Value = rowsForCalcluate4Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["Total"].Value));
                            row4.Cells["TotalPlan"].Value = rowsForCalcluate4Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["TotalPlan"].Value));
                            row4.Cells["TotalPlanCel"].Value = GlobalUtils.TryParseDecimal(row4.Cells["Total"].Value) + GlobalUtils.TryParseDecimal(row4.Cells["TotalPlan"].Value); // Итого цел + план
                        }
                        catch (Exception ex)
                        { }
                    }
                    /// ПОДСЧЕТ ИТОГОВЫХ СТРОК ДЛЯ ТАБЛИЦЫ 11 
                    if (Report.IdType == "PG_Q" && GetCurrentTheme() == "Таблица 11")
                    {
                        try
                        {
                            string[] rowFor1_1_3Row = { "1.1.3.1", "1.1.3.2" };
                            var row1_1_3 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "1.1.3");
                            var rowsForCalcluate1_1_3Total = Dgv.Rows.Cast<DataGridViewRow>().Where(x => rowFor1_1_3Row.Contains(x.Cells[1].Value.ToString()));
                            row1_1_3.Cells["Total"].Value = rowsForCalcluate1_1_3Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["Total"].Value));
                            row1_1_3.Cells["TotalPlan"].Value = rowsForCalcluate1_1_3Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["TotalPlan"].Value));
                            row1_1_3.Cells["TotalPlanCel"].Value = GlobalUtils.TryParseDecimal(row1_1_3.Cells["Total"].Value) + GlobalUtils.TryParseDecimal(row1_1_3.Cells["TotalPlan"].Value); // Итого цел + план
                        }
                        catch (Exception ex)
                        { }
                        try
                        {
                            string[] rowFor1_1Row = { "1.1.1", "1.1.2", "1.1.3" };
                            var row1_1 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "1.1");
                            var rowsForCalcluate1_1Total = Dgv.Rows.Cast<DataGridViewRow>().Where(x => rowFor1_1Row.Contains(x.Cells[1].Value.ToString()));
                            row1_1.Cells["Total"].Value = rowsForCalcluate1_1Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["Total"].Value));
                            row1_1.Cells["TotalPlan"].Value = rowsForCalcluate1_1Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["TotalPlan"].Value));
                            row1_1.Cells["TotalPlanCel"].Value = GlobalUtils.TryParseDecimal(row1_1.Cells["Total"].Value) + GlobalUtils.TryParseDecimal(row1_1.Cells["TotalPlan"].Value); // Итого цел + план
                        }
                        catch (Exception ex)
                        { }
                        try
                        {
                            string[] rowFor3Row = { "3.1", "3.2", "3.3" };
                            var row3 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "3");
                            var rowsForCalcluate3Total = Dgv.Rows.Cast<DataGridViewRow>().Where(x => rowFor3Row.Contains(x.Cells[1].Value.ToString()));
                            row3.Cells["Total"].Value = rowsForCalcluate3Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["Total"].Value));
                            row3.Cells["TotalPlan"].Value = rowsForCalcluate3Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["TotalPlan"].Value));
                            row3.Cells["TotalPlanCel"].Value = GlobalUtils.TryParseDecimal(row3.Cells["Total"].Value) + GlobalUtils.TryParseDecimal(row3.Cells["TotalPlan"].Value); // Итого цел + план
                        }
                        catch (Exception ex)
                        { }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

            }

        }

        public void MapReportFromDgv(string form)
        {
            if (Dgv.CurrentCell != null && Dgv.CurrentCell.IsInEditMode)
            {
                Dgv.EndEdit();
            }

            FillReport(form);
        }

        public void CreateReportForm(string form)
        {
            Dgv.AutoGenerateColumns = false;
            Dgv.AllowUserToAddRows = false;
            Dgv.AutoSize = true;
            Dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            Dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            Dgv.Columns.Clear();
            Dgv.Rows.Clear();

            if (form == "Свод")
                form = "ОПЭД";


            var table = ThemeTextData.Tables_fromxml
                .Where(x => x.TableName_fromxml == form)
                .SelectMany(x => x.Rows_fromxml)
                .ToList();



            CreateDgvForForm(form, table);

            if (Report.IdType == "PG" || Report.IdType == "PG_Q" || Report.IdType == "foped" || Report.IdType == "Zpz_Q" || Report.IdType == "Zpz" || Report.IdType == "ZpzLethal" || Report.IdType == "Zpz_Q2025" || Report.IdType == "Zpz2025" || Report.IdType == "ZpzL2025"/**|| Report.IdType == "Zpz10"**/)
                CreateTotalColumn();

        }

        public void SetReadonlyForDgv(bool isReadonly) => Dgv.ReadOnly = isReadonly;

        public string GetCurrentTheme() => Cmb.Text;

        private List<KmsReportDictionary> GetForms()
        {
            try
            {
                var result = ThemeTextData.Tables_fromxml
               .Select(x => new KmsReportDictionary { Key = x.TableName_fromxml, Value = x.TableDescription_fromxml })
               .ToList();

                if (IdReportType == "foped")
                    result.Add(new KmsReportDictionary { Key = "Свод", Value = "За год" });

                return result;



            }
            catch (Exception ex)
            {
                // Логирование ошибки для отладки
                Console.WriteLine($"Ошибка при получении форм: {ex.Message}");
                // Можно использовать специализированные библиотеки для логирования, такие как NLog или Serilog
                // Для примера просто выбрасываем исключение с дополнительной информацией
                throw new ApplicationException("Ошибка при получении форм из XML", ex);
            }


        }


        public string GetReportInfo()
        {
            string info = $"{SmallName}; Период: {Report.Yymm}; " + Environment.NewLine;

            if (Report.IdEmployee != 0)
            {
                info += $"Дата создания: {Report.Created.ToShortDateString()} ";
                info += $"Пользователь: {GetUser(Report.IdEmployee)}; " + Environment.NewLine;
            }

            if (Report.IdEmployeeUpd != 0 && Report.Updated != null)
            {
                info += $"Дата обновления: {Report.Updated.Value.ToShortDateString()} ";
                info += $"Пользователь: {GetUser(Report.IdEmployeeUpd)}; " + Environment.NewLine;
            }

            info += "Наличие скана: " + (!string.IsNullOrEmpty(Report.Scan) ? "Да; " : "Нет; ") + Environment.NewLine;
            if (Report.UserToCo != 0 && Report.DateToCo != null)
            {
                info += $"Дата направления в ЦО: {Report.DateToCo.Value.ToShortDateString()} ";
                info += $"Пользователь: {GetUser(Report.UserToCo)}; " + Environment.NewLine;
            }

            if (Report.IdType == "PG" || Report.IdType == "Zpz" || Report.IdType == "Zpz2025")
            {
                info += "Данные вносились вручную: " + (Report.DataSource.ToString() != "Handle" ? "Нет; " : "Да; ") + Environment.NewLine;
            }

            if (Report.RefuseUser != 0 && Report.RefuseDate != null)
            {
                info += $"Дата возврата отчета на доработку: {Report.RefuseDate.Value.ToShortDateString()} ";
                info += $"Пользователь: {GetUser(Report.RefuseUser)}; " + Environment.NewLine;
            }

            if (Report.UserSubmit != 0 && Report.DateIsDone != null)
            {
                info += $"Дата утверждения: {Report.DateIsDone.Value.ToShortDateString()} ";
                info += $"Пользователь: {GetUser(Report.UserSubmit)}; " + Environment.NewLine;
            }

            return info;
        }

        public void SaveScan(string inUri, int num)
        {
            try
            {
                var request = new SaveScanRequest
                {
                    Body = new SaveScanRequestBody
                    {
                        idReport = Report.IdFlow,
                        idUser = CurrentUser.IdUser,
                        uri = inUri,
                        num = num
                    }
                };
                Client.SaveScan(request);
            }
            catch (Exception ex)
            {
                log.Error(ex, "Ошибка сохранения скана в БД");
                throw;
            }
        }

        public void DeleteScan(int num)
        {
            try
            {
                Client.DeleteScan(Report.IdFlow, CurrentUser.IdUser, num = num);
            }
            catch (Exception ex)
            {
                log.Error(ex, "Ошибка удаления скана в БД");
                throw;
            }
        }


        public void ChangeDataSource(DataSource datasource)
        {
            try
            {
                Client.ChangeDataSource(Report.IdFlow, CurrentUser.IdUser, datasource);
                Dgv.ReadOnly = datasource == DataSource.New;
            }
            catch (Exception ex)
            {
                log.Error(ex, "Ошибка при попытке загрузить отчет из Excel");
                throw;
            }
        }

        public void ChangeStatus(ReportStatus status)
        {
            try
            {
                Client.ChangeStatus(Report.IdFlow, CurrentUser.IdUser, status);

                Dgv.ReadOnly = status == ReportStatus.Done || status == ReportStatus.Submit;
            }
            catch (Exception ex)
            {
                log.Error(ex, "Ошибка утверждения филиалом отчета");
                throw;
            }
        }

        public void DeserializeReport(string yymm)
        {
            var binFormat = new BinaryFormatter();
            var filename = GlobalUtils.GetSerializeName(SerializeName, yymm);
            if (File.Exists(filename))
            {
                using Stream fStream = File.OpenRead(filename);
                Report = (TR)binFormat.Deserialize(fStream);
            }
        }

        public void Serialize(string yymm)
        {
            var filename = GlobalUtils.GetSerializeName(SerializeName, yymm);
            var binFormat = new BinaryFormatter();
            using Stream fStream = new FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.None);
            binFormat.Serialize(fStream, Report);
        }

        AbstractReport IReportProcessor.Report {
            get => Report;
            set {
                Report = (TR)value;
                Report.IdType = IdReportType;
            }
        }

        Color IReportProcessor.ColorReport {
            get => ColorReport;
            set => ColorReport = value;
        }

        bool IReportProcessor.HasReport {
            get => HasReport;
            set => HasReport = value;
        }

        string IReportProcessor.OldTheme {
            get => OldTheme;
            set => OldTheme = value;
        }

        public List<KmsReportDictionary> ThemesList {
            get => GetForms();
        }

        TabPage IReportProcessor.Page {
            get => Page;
        }

        string IReportProcessor.SmallName {
            get => SmallName;
        }

        string IReportProcessor.FilialName {
            get => FilialName;
        }

        string IReportProcessor.FilialCode {
            get => FilialCode;
            set {
                FilialCode = value;
                if (!CurrentUser.IsMain || Report.Status != ReportStatus.Done)
                {
                    FilialName = CurrentUser.Filial;
                }
                else
                {
                    FilialName = CurrentUser.Regions.Single(x => x.Key == value).ForeignKey;
                }
            }
        }

        protected abstract void FillReport(string form);

        protected abstract void CreateDgvForForm(string form, List<TemplateRow> table);

        protected bool IsNotNeedFillRow(string form, string rowNum) =>
            ThemeTextData.Tables_fromxml
                .Where(x => x.TableName_fromxml == form)
                .SelectMany(x => x.Rows_fromxml)
                .Single(x => x.RowNum_fromxml == rowNum)
                .Exclusion_fromxml;

        protected TemplateForm ReadTemplateXml(string filename)
        {
            try
            {
                var xmlDoc = XDocument.Load(filename);
                var xmlSerializer = new XmlSerializer(typeof(TemplateForm));
                if (xmlDoc.Root != null)
                {
                    using var reader = xmlDoc.Root.CreateReader();
                    return (TemplateForm)xmlSerializer.Deserialize(reader);
                }
            }
            catch (Exception ex)
            {
                log.Error($"Ошибка получения данных из шаблона для формы: {ex}");
            }

            return null;
        }

        private string GetUser(int key)
        {
            return CurrentUser.Users.SingleOrDefault(x => x.Key == key.ToString())?.Value ?? "Ошибка получения имени";
        }

        public void CallculateCells()
        {
            this.CalculateCells();
        }
    }
}