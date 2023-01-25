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
            Dgv.Columns.Add("Total", "Итого");
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
        }

        public void SetTotalColumn()
        {
            try
            {
                int columnCount = Dgv.Columns.Count - 1;
                for (int row = 0; row < Dgv.Rows.Count; row++)
                {
                    Dgv.Rows[row].Cells[columnCount].Value = 0;
                    decimal valueCel = 0.00M;
                    decimal valuePlan = 0.00M;

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
                                    valueCel += GlobalUtils.TryParseDecimal(Dgv.Rows[row].Cells[cell].Value);
                                }
                                else if (Dgv.Rows[row].Cells[cell].ColumnIndex == 8 || Dgv.Rows[row].Cells[cell].ColumnIndex == 9 || Dgv.Rows[row].Cells[cell].ColumnIndex == 10 || Dgv.Rows[row].Cells[cell].ColumnIndex == 12)
                                {
                                    valuePlan += GlobalUtils.TryParseDecimal(Dgv.Rows[row].Cells[cell].Value);
                                }
                                /// НА ДАННОМ ЭТАПЕ ПРОБЕГАЕМ ДЛЯ КАЖДОЙ СТРОКИ ОТЧЕТА ПО ЯЧЕЙКАМ и СУММИРУЕМ ЗНАЧЕНИЯ ДЛЯ ЦЕЛЕВЫХ (2-3-4-6) И ПЛАНОВЫХ (8-9-10-12), ПОКА ПРОСТО В ПЕРЕМЕННЫЕ valueCel и valuePlan /// 
                            }

                            else
                            {
                                if ((GetCurrentTheme() == "Таблица 1Л" || GetCurrentTheme() == "Таблица 2Л"))
                                {
                                    if (Dgv.Rows[row].Cells[cell].ColumnIndex != 6)
                                    {
                                        valueCel += GlobalUtils.TryParseDecimal(Dgv.Rows[row].Cells[cell].Value);

                                    }
                                    // Потребовали, чтобы в 6й колонке была сумма 2-5. Попробуем плюсануть тут.
                                    else if (GetCurrentTheme() == "Таблица 1Л" && row != Dgv.Rows.Count - 1 && row != Dgv.Rows.Count - 2)
                                    { Dgv.Rows[row].Cells[cell].Value = valueCel; };
                                }
                                else if (GetCurrentTheme() == "Таблица 12")
                                {
                                    if (Dgv.Rows[row].Cells[cell].ColumnIndex != 3)
                                    {
                                        valueCel += GlobalUtils.TryParseDecimal(Dgv.Rows[row].Cells[cell].Value);
                                    }
                                }
                                else
                                {
                                    valueCel += GlobalUtils.TryParseDecimal(Dgv.Rows[row].Cells[cell].Value);
                                }
                            }
                        }
                    }

                    //Тот, кто это видит прошу меня простить))
                    if ((Report.IdType == "PG" || Report.IdType == "PG_Q") && (GetCurrentTheme() == "Таблица 6" || GetCurrentTheme() == "Таблица 8" || GetCurrentTheme() == "Таблица 10"))
                    {
                        Dgv.Rows[row].Cells["Total"].Value = valueCel; //Целевые
                        Dgv.Rows[row].Cells["TotalPlan"].Value = valuePlan; // Плановые
                        Dgv.Rows[row].Cells["TotalPlanCel"].Value = valuePlan + valueCel; // Итого цел + план
                        // Пишем в DGV значения и сумму
                    }
                    else
                    {
                        Dgv.Rows[row].Cells[columnCount].Value = valueCel; //Целевые
                    }
                }

                string[] rowFor6Row = { "6", "6.1", "6.2", "6.3", "6.4", "6.5", "6.6", "6.7", "6.8", "6.9", "6.10" };
                if ((Report.IdType == "PG" || Report.IdType == "PG_Q") && (GetCurrentTheme() == "Таблица 6" || GetCurrentTheme() == "Таблица 8" || GetCurrentTheme() == "Таблица 5" || GetCurrentTheme() == "Таблица 10"))
                {
                    var row6 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "6");
                    if (row6 != null)
                    {
                        var rowsForCalcluate6Total = Dgv.Rows.Cast<DataGridViewRow>().Where(x => rowFor6Row.Contains(x.Cells[1].Value.ToString()));

                        row6.Cells["Total"].Value = rowsForCalcluate6Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["Total"].Value));
                        row6.Cells["TotalPlan"].Value = rowsForCalcluate6Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["TotalPlan"].Value));

                        row6.Cells["TotalPlanCel"].Value = GlobalUtils.TryParseDecimal(row6.Cells["Total"].Value) + GlobalUtils.TryParseDecimal(row6.Cells["TotalPlan"].Value); // Итого цел + план
                    }

                    //    if (Report.IdType == "PG" && GetCurrentTheme() == "Таблица 6")
                    //    {
                    //        string[] rowFor5Row = { "5", "5.3", "5.4", "5.5", "5.6", "5.7", "5.8" };
                    //        foreach (string rowi in rowFor5Row)
                    //        {
                    //            var row5 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == rowi);
                    //            var rowsForCalcluate5Total = Dgv.Rows.Cast<DataGridViewRow>().Where(x => rowFor5Row.Contains(x.Cells[1].Value.ToString()));
                    //            row5.Cells["Total"].Value = rowsForCalcluate5Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["Total"].Value));
                    //            row5.Cells["TotalPlan"].Value = rowsForCalcluate5Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["TotalPlan"].Value));

                    //            row5.Cells["TotalPlanCel"].Value = GlobalUtils.TryParseDecimal(row5.Cells["Total"].Value) + GlobalUtils.TryParseDecimal(row5.Cells["TotalPlan"].Value); // Итого цел + план
                    //        }
                    //    }
                    //    else if (Report.IdType == "PG_Q" && GetCurrentTheme() == "Таблица 6")
                    //    {
                    //        string[] rowFor5Row = { "5", "5.1", "5.2", "5.3", "5.4", "5.5", "5.6", "5.7", "5.8" };
                    //        var row5 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "5");
                    //        var rowsForCalcluate5Total = Dgv.Rows.Cast<DataGridViewRow>().Where(x => rowFor5Row.Contains(x.Cells[1].Value.ToString()));
                    //        row5.Cells["Total"].Value = rowsForCalcluate5Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["Total"].Value));
                    //        row5.Cells["TotalPlan"].Value = rowsForCalcluate5Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["TotalPlan"].Value));
                    //        row5.Cells["TotalPlanCel"].Value = GlobalUtils.TryParseDecimal(row5.Cells["Total"].Value) + GlobalUtils.TryParseDecimal(row5.Cells["TotalPlan"].Value); // Итого цел + план

                    //    }

                    //    if (Report.IdType == "PG_Q" && GetCurrentTheme() == "Таблица 5")
                    //    {
                    //        string[] rowFor4Row = { "4.1", "4.2", "4.3", "4.4", "4.5", "4.6" };
                    //        var row4 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "4");
                    //        var rowsForCalcluate5Total = Dgv.Rows.Cast<DataGridViewRow>().Where(x => rowFor4Row.Contains(x.Cells[1].Value.ToString()));
                    //        row4.Cells["Total"].Value = rowsForCalcluate5Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["Total"].Value));
                    //        row4.Cells["TotalPlan"].Value = rowsForCalcluate5Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["TotalPlan"].Value));

                    //        row4.Cells["TotalPlanCel"].Value = GlobalUtils.TryParseDecimal(row4.Cells["Total"].Value) + GlobalUtils.TryParseDecimal(row4.Cells["TotalPlan"].Value); // Итого цел + план

                    //    }

                    if ((Report.IdType == "PG_Q" || Report.IdType == "PG") && GetCurrentTheme() == "Таблица 10")
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

                        finally
                        {
                            string[] rowFor5Row = { "5.1", "5.2", "5.3", "5.4", "5.5", "5.6", "5.7", "5.8" };
                            var row5 = Dgv.Rows.Cast<DataGridViewRow>().FirstOrDefault(x => x.Cells[1].Value.ToString() == "5");
                            var rowsForCalcluate5Total = Dgv.Rows.Cast<DataGridViewRow>().Where(x => rowFor5Row.Contains(x.Cells[1].Value.ToString()));
                            row5.Cells["Total"].Value = rowsForCalcluate5Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["Total"].Value));
                            row5.Cells["TotalPlan"].Value = rowsForCalcluate5Total.Sum(x => GlobalUtils.TryParseDecimal(x.Cells["TotalPlan"].Value));

                            row5.Cells["TotalPlanCel"].Value = GlobalUtils.TryParseDecimal(row5.Cells["Total"].Value) + GlobalUtils.TryParseDecimal(row5.Cells["TotalPlan"].Value); // Итого цел + план
                        }
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
            Dgv.AutoSize = false;
            Dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            Dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            Dgv.Columns.Clear();
            Dgv.Rows.Clear();

            if (form == "Свод")
                form = "ОПЭД";


            var table = ThemeTextData.tables
                .Where(x => x.Name == form)
                .SelectMany(x => x.Rows)
                .ToList();



            CreateDgvForForm(form, table);

            if (Report.IdType == "PG" || Report.IdType == "PG_Q" || Report.IdType == "foped" || Report.IdType == "Zpz_Q" || Report.IdType == "Zpz")
                CreateTotalColumn();

        }

        public void SetReadonlyForDgv(bool isReadonly) => Dgv.ReadOnly = isReadonly;

        public string GetCurrentTheme() => Cmb.Text;

        private List<KmsReportDictionary> GetForms()
        {
            try
            {
                var result = ThemeTextData.tables
               .Select(x => new KmsReportDictionary { Key = x.Name, Value = x.TableDescription })
               .ToList();

                if (IdReportType == "foped")
                    result.Add(new KmsReportDictionary { Key = "Свод", Value = "За год" });

                return result;



            }
            catch (Exception ex)
            {

                throw;
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

            if (Report.IdType == "PG" || Report.IdType == "Zpz")
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
            ThemeTextData.tables
                .Where(x => x.Name == form)
                .SelectMany(x => x.Rows)
                .Single(x => x.Num == rowNum)
                .Exclusion;

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