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
using KmsReportClient.Model.XML;
using KmsReportClient.Support;
using NLog;

namespace KmsReportClient.Report.Basic
{
    public class ReportVaccinationProccesor : AbstractReportProcessor<ReportVaccination>
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();
        StackedHeaderDecorator DgvRender;

        private int[] _columnsReadOnly = { 1, 2, 8, 14, 15, 21 };

        private readonly List<string> _columns = new List<string>
        {
            ";ВСЕГО",

            ";Мужчины;Всего",
            ";Мужчины;Из них 18-39 лет",
            ";Мужчины;Из них 40-59 лет",
            ";Мужчины;Из них 60-65 лет",
            ";Мужчины;Из них 66-74 лет",
            ";Мужчины;Из них 75 лет и старше",

            ";Женщины;Всего",
            ";Женщины;Из них 18-39 лет",
            ";Женщины;Из них 40-54 лет",
            ";Женщины;Из них 55-65 лет",
            ";Женщины;Из них 66-74 лет",
            ";Женщины;Из них 75 лет и старше",

        };



        public ReportVaccinationProccesor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
       base(inClient, dgv, cmb, txtb, page,
           XmlFormTemplate.Vac.GetDescription(),
           Log,
           ReportGlobalConst.ReportVac,
           reportsDictionary)
        {
            DgvRender = new StackedHeaderDecorator(Dgv);
            InitReport();

        }


        public override AbstractReport CollectReportFromWs(string yymm)
        {
            var request = new GetReportRequest
            {
                Body = new GetReportRequestBody
                {
                    filialCode = FilialCode,
                    yymm = yymm,
                    reportType = ReportType.Vac
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as ReportVaccination;
        }
        public override void FillDataGridView(string form)
        {

            if (Report == null)
            {
                return;
            }


            Dgv.Rows[0].Cells[0].Value = Report.Id;

            Dgv.Rows[0].Cells[3].Value = Report.M18_39;
            Dgv.Rows[0].Cells[4].Value = Report.M40_59;
            Dgv.Rows[0].Cells[5].Value = Report.M60_65;
            Dgv.Rows[0].Cells[6].Value = Report.M66_74;
            Dgv.Rows[0].Cells[7].Value = Report.M75_More;


            Dgv.Rows[0].Cells[9].Value = Report.W18_39;
            Dgv.Rows[0].Cells[10].Value = Report.W40_54;
            Dgv.Rows[0].Cells[11].Value = Report.W55_65;
            Dgv.Rows[0].Cells[12].Value = Report.W66_74;
            Dgv.Rows[0].Cells[13].Value = Report.W75_More;


            GetYearData();
            SetFormulaMonth();
           
        }

        public override void SaveReportDataSourceExcel()
        { }
        public override void SaveReportDataSourceHandle()
        { }

        public void SetFormulaMonth()
        {
            try
            {
                //Всего мужчины
                Dgv.Rows[0].Cells[2].Value =    GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[3].Value) +
                                                GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[4].Value) +
                                                GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[5].Value) +
                                                GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[6].Value) + 
                                                GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[7].Value);

                //Всего мужчины
                Dgv.Rows[0].Cells[8].Value = GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[9].Value) +
                                                GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[10].Value) +
                                                GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[11].Value) +
                                                GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[12].Value) +
                                                GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[13].Value);
                //Всего вообще
                Dgv.Rows[0].Cells[1].Value = GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[2].Value) + GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[8].Value);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
       
        }

        public void SetFormulaYear()
        {
            //Всего мужчины
            Dgv.Rows[0].Cells[15].Value = GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[16].Value) +
                                            GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[17].Value) +
                                            GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[18].Value) +
                                            GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[19].Value) +
                                            GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[20].Value);

            //Всего мужчины
            Dgv.Rows[0].Cells[21].Value = GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[22].Value) +
                                            GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[23].Value) +
                                            GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[24].Value) +
                                            GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[25].Value) +
                                            GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[26].Value);
            //Всего вообще
            Dgv.Rows[0].Cells[14].Value = GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[15].Value) + GlobalUtils.TryParseInt(Dgv.Rows[0].Cells[21].Value);
        }
        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource)
        {

        }
        public override void InitReport()
        {
            Report = new ReportVaccination { IdType = IdReportType };

        }
        public override bool IsVisibleBtnDownloadExcel() => true;

        public override bool IsVisibleBtnHandle() => false;


        public override void MapForAutoFill(AbstractReport report)
        {

            var inReport = report as ReportVaccination;
            Report.Id = inReport.Id;
            Report.IdReportData = inReport.IdReportData;

            Report.M18_39 = inReport.M18_39;
            Report.M40_59 = inReport.M40_59;
            Report.M60_65 = inReport.M60_65;
            Report.M66_74 = inReport.M66_74;
            Report.M75_More = inReport.M75_More;

            Report.W18_39 = inReport.W18_39;
            Report.W40_54 = inReport.W40_54;
            Report.W55_65 = inReport.W55_65;
            Report.W66_74 = inReport.W66_74;
            Report.W75_More = inReport.W75_More;


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
                    reportType = ReportType.Vac
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportVaccination;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;


            GetYearData();

        }


        public void GetYearData()
        {

            var yearThemeData = Client.GetVacYearData(new GetVacYearDataRequest(new GetVacYearDataRequestBody
            {
                fillial = FilialCode,
                yymm = Report.Yymm
            })).Body.GetVacYearDataResult;

            if (yearThemeData != null)
            {

                Dgv.Rows[0].Cells[16].Value = yearThemeData.M18_39;
                Dgv.Rows[0].Cells[17].Value = yearThemeData.M40_59;
                Dgv.Rows[0].Cells[18].Value = yearThemeData.M60_65;
                Dgv.Rows[0].Cells[19].Value = yearThemeData.M66_74;
                Dgv.Rows[0].Cells[20].Value = yearThemeData.M75_More;


                Dgv.Rows[0].Cells[22].Value = yearThemeData.W18_39;
                Dgv.Rows[0].Cells[23].Value = yearThemeData.W40_54;
                Dgv.Rows[0].Cells[24].Value = yearThemeData.W55_65;
                Dgv.Rows[0].Cells[25].Value = yearThemeData.W66_74;
                Dgv.Rows[0].Cells[26].Value = yearThemeData.W75_More;

                SetFormulaYear();
            }
        }

        public override void ToExcel(string filename, string filialName)
        {
            var excel = new ExcelVacCreator(filename, ExcelForm.Vac, Report.Yymm, filialName, Client, FilialCode);
            excel.CreateReport(Report, null);
        }
        public override string ValidReport()
        {
            return "";
        }
        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            Dgv.AllowUserToAddRows = false;
            Dgv.ColumnHeadersVisible = true;

            Dgv.Columns.Clear();
            Dgv.Rows.Clear();

            //Добавляем столбец с ID
            Dgv.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Id",
                DataPropertyName = "Indicator",
                Name = "Indicator",
                SortMode = DataGridViewColumnSortMode.NotSortable,
                DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Azure },
                Visible = false
            });



            //Создание столбцов текущего месяца

            string reportMonth = YymmUtils.GetMonth(Report.Yymm.Substring(2));
            string prefixHeader = $"Текущий месяц;{reportMonth};Проинформировано о вакцинации";
            foreach (var clmn in _columns)
            {
                var column = new DataGridViewTextBoxColumn
                {
                    HeaderText = prefixHeader + clmn,
                    DataPropertyName = "Indicator",
                    Name = "Indicator",
                    SortMode = DataGridViewColumnSortMode.NotSortable,
                    DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Azure }
                };

                Dgv.Columns.Add(column);
            }

            int year = YymmUtils.ConvertYymmToDate(Report.Yymm).Year;
            prefixHeader = $"Сначала года;{year};Проинформировано о вакцинации";
            // Создание столбцов для года
            foreach (var clmn in _columns)
            {

                var column = new DataGridViewTextBoxColumn
                {
                    HeaderText = clmn.ToLower() == "всего" ? clmn : prefixHeader + clmn,
                    DataPropertyName = "Indicator",
                    Name = "Indicator",
                    SortMode = DataGridViewColumnSortMode.NotSortable,
                    DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Azure }
                };

                Dgv.Columns.Add(column);
            }

            Dgv.Rows.Add();
         
            foreach (var c in _columnsReadOnly)
            {
                Dgv.Columns[c].ReadOnly = true;
                Dgv.Columns[c].DefaultCellStyle.BackColor = Color.LightGray;

            }


        }
        protected override void FillReport(string form)
        {

            var row = Dgv.Rows[0];
            Report.Id = GlobalUtils.TryParseInt(row.Cells[0].Value);

            Report.M18_39 = GlobalUtils.TryParseInt(row.Cells[3].Value);
            Report.M40_59 = GlobalUtils.TryParseInt(row.Cells[4].Value);
            Report.M60_65 = GlobalUtils.TryParseInt(row.Cells[5].Value);
            Report.M66_74 = GlobalUtils.TryParseInt(row.Cells[6].Value);
            Report.M75_More = GlobalUtils.TryParseInt(row.Cells[7].Value);

            Report.W18_39 = GlobalUtils.TryParseInt(row.Cells[9].Value);
            Report.W40_54 = GlobalUtils.TryParseInt(row.Cells[10].Value);
            Report.W55_65 = GlobalUtils.TryParseInt(row.Cells[11].Value);
            Report.W66_74 = GlobalUtils.TryParseInt(row.Cells[12].Value);
            Report.W75_More = GlobalUtils.TryParseInt(row.Cells[13].Value);

        }
    }
}
