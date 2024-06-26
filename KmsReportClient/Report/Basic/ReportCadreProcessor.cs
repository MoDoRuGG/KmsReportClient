﻿using System;
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
    public class ReportCadreProcessor : AbstractReportProcessor<ReportCadre>
    {
        StackedHeaderDecorator DgvRender;

        private readonly List<string> zpz_ekmp = new List<string>
        {
            "Численность (шт. ед.), без спец-ов;по штату;3",
            "Численность (шт. ед.), без спец-ов;факт;4",
            "Численность (шт. ед.), без спец-ов;вакансии;5",
            "в том числе:;Руководитель (шт. ед.);по штату;6",
            "в том числе:;Руководитель (шт. ед.);факт;7",
            "в том числе:;Руководитель (шт. ед.);вакансии;8",
            "в том числе:;Заместитель руководителя (шт. ед.);по штату;9",
            "в том числе:;Заместитель руководителя (шт. ед.);факт;10",
            "в том числе:;Заместитель руководителя (шт. ед.);вакансии;11",
            "в том числе:;Врачи-эксперты (исключая рук-во) (шт. ед.);по штату;12",
            "в том числе:;Врачи-эксперты (исключая рук-во) (шт. ед.);факт;13",
            "в том числе:;Врачи-эксперты (исключая рук-во) (шт. ед.);вакансии;14",
            "в том числе:;Общее кол-во врачей-экспертов (из гр.13)(физ. лица); Х > 1,0;15",
            "в том числе:;Общее кол-во врачей-экспертов (из гр.13)(физ. лица); Х = 1,0;16",
            "в том числе:;Общее кол-во врачей-экспертов (из гр.13)(физ. лица); Х < 1,0;Х = 0,75;17",
            "в том числе:;Общее кол-во врачей-экспертов (из гр.13)(физ. лица); Х < 1,0;Х = 0,5 (0,6);18",
            "в том числе:;Общее кол-во врачей-экспертов (из гр.13)(физ. лица); Х < 1,0;Х = 0,25;19",
            "в том числе:;Общее кол-во врачей-экспертов (из гр.13)(физ. лица); Х < 1,0;X <= 0,1;20",
            "в том числе:;Врачи-эксперты качества МП, входящие в реестр врачей-экспертов ФОМС (из общего кол-ва врачей-экспертов гр. (15, 16, 17, 18, 19, 20) (физ. лица); Х > 1,0;21",
            "в том числе:;Врачи-эксперты качества МП, входящие в реестр врачей-экспертов ФОМС (из общего кол-ва врачей-экспертов гр. (15, 16, 17, 18, 19, 20) (физ. лица); Х = 1,0;22",
            "в том числе:;Врачи-эксперты качества МП, входящие в реестр врачей-экспертов ФОМС (из общего кол-ва врачей-экспертов гр. (15, 16, 17, 18, 19, 20) (физ. лица); Х < 1,0;Х = 0,75;23",
            "в том числе:;Врачи-эксперты качества МП, входящие в реестр врачей-экспертов ФОМС (из общего кол-ва врачей-экспертов гр. (15, 16, 17, 18, 19, 20) (физ. лица); Х < 1,0;Х = 0,5 (0,6);24",
            "в том числе:;Врачи-эксперты качества МП, входящие в реестр врачей-экспертов ФОМС (из общего кол-ва врачей-экспертов гр. (15, 16, 17, 18, 19, 20) (физ. лица); Х < 1,0;Х = 0,25;25",
            "в том числе:;Врачи-эксперты качества МП, входящие в реестр врачей-экспертов ФОМС (из общего кол-ва врачей-экспертов гр. (15, 16, 17, 18, 19, 20) (физ. лица); Х < 1,0;X <= 0,1;26",
            "в том числе:;Специалисты;по штату;27",
            "в том числе:;Специалисты;факт;28",
            "в том числе:;Специалисты;вакансии;29",
        };

        private readonly List<string> oi_zpz = new List<string>
        {
            "Численность (шт. ед.), без спец-ов;по штату;3",
            "Численность (шт. ед.), без спец-ов;факт;4",
            "Численность (шт. ед.), без спец-ов;вакансии;5",
            "в том числе:;Руководитель (шт. ед.);по штату;6",
            "в том числе:;Руководитель (шт. ед.);факт;7",
            "в том числе:;Руководитель (шт. ед.);вакансии;8",
            "в том числе:;Заместитель руководителя (шт. ед.);по штату;9",
            "в том числе:;Заместитель руководителя (шт. ед.);факт;10",
            "в том числе:;Заместитель руководителя (шт. ед.);вакансии;11",
            "в том числе:;Врачи-эксперты (исключая рук-во) (шт. ед.);по штату;12",
            "в том числе:;Врачи-эксперты (исключая рук-во) (шт. ед.);факт;13",
            "в том числе:;Врачи-эксперты (исключая рук-во) (шт. ед.);вакансии;14",
            "в том числе:;Общее кол-во врачей-экспертов (из гр.13)(физ. лица); Х > 1,0;15",
            "в том числе:;Общее кол-во врачей-экспертов (из гр.13)(физ. лица); Х = 1,0;16",
            "в том числе:;Общее кол-во врачей-экспертов (из гр.13)(физ. лица); Х < 1,0;Х = 0,75;17",
            "в том числе:;Общее кол-во врачей-экспертов (из гр.13)(физ. лица); Х < 1,0;Х = 0,5 (0,6);18",
            "в том числе:;Общее кол-во врачей-экспертов (из гр.13)(физ. лица); Х < 1,0;Х = 0,25;19",
            "в том числе:;Общее кол-во врачей-экспертов (из гр.13)(физ. лица); Х < 1,0;X <= 0,1;20",
            "в том числе:;Врачи-эксперты качества МП, входящие в реестр врачей-экспертов ФОМС (из общего кол-ва врачей-экспертов гр. (15, 16, 17, 18, 19, 20) (физ. лица); Х > 1,0;21",
            "в том числе:;Врачи-эксперты качества МП, входящие в реестр врачей-экспертов ФОМС (из общего кол-ва врачей-экспертов гр. (15, 16, 17, 18, 19, 20) (физ. лица); Х = 1,0;22",
            "в том числе:;Врачи-эксперты качества МП, входящие в реестр врачей-экспертов ФОМС (из общего кол-ва врачей-экспертов гр. (15, 16, 17, 18, 19, 20) (физ. лица); Х < 1,0;Х = 0,75;23",
            "в том числе:;Врачи-эксперты качества МП, входящие в реестр врачей-экспертов ФОМС (из общего кол-ва врачей-экспертов гр. (15, 16, 17, 18, 19, 20) (физ. лица); Х < 1,0;Х = 0,5 (0,6);24",
            "в том числе:;Врачи-эксперты качества МП, входящие в реестр врачей-экспертов ФОМС (из общего кол-ва врачей-экспертов гр. (15, 16, 17, 18, 19, 20) (физ. лица); Х < 1,0;Х = 0,25;25",
            "в том числе:;Врачи-эксперты качества МП, входящие в реестр врачей-экспертов ФОМС (из общего кол-ва врачей-экспертов гр. (15, 16, 17, 18, 19, 20) (физ. лица); Х < 1,0;X <= 0,1;26",
            "в том числе:;Специалисты;по штату;27",
            "в том числе:;Специалисты;факт;28",
            "в том числе:;Специалисты;вакансии;29",
        };

        private static readonly Logger Log = LogManager.GetCurrentClassLogger();
        public ReportCadreProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
        base(inClient, dgv, cmb, txtb, page,
            XmlFormTemplate.Cadre.GetDescription(),
            Log,
            ReportGlobalConst.ReportCadre,
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
                    reportType = ReportType.Cadre
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as ReportCadre;

        }
        public override void FillDataGridView(string form)
        {

            var reportCadre = Report.ReportDataList.FirstOrDefault(x => x.Theme == form);
            if (reportCadre == null)
            {
                return;
            }

            if (reportCadre.Data != null)
            {
                Dgv.Rows[0].Cells[0].Value = reportCadre.Data.count_itog_state;
                Dgv.Rows[0].Cells[1].Value = reportCadre.Data.count_itog_fact;
                Dgv.Rows[0].Cells[2].Value = reportCadre.Data.count_itog_vacancy;
                Dgv.Rows[0].Cells[3].Value = reportCadre.Data.count_leader_state;
                Dgv.Rows[0].Cells[4].Value = reportCadre.Data.count_leader_fact;
                Dgv.Rows[0].Cells[5].Value = reportCadre.Data.count_leader_vacancy;
                Dgv.Rows[0].Cells[6].Value = reportCadre.Data.count_deputy_leader_state;
                Dgv.Rows[0].Cells[7].Value = reportCadre.Data.count_deputy_leader_fact;
                Dgv.Rows[0].Cells[8].Value = reportCadre.Data.count_deputy_leader_vacancy;
                Dgv.Rows[0].Cells[9].Value = reportCadre.Data.count_expert_doctor_state;
                Dgv.Rows[0].Cells[10].Value = reportCadre.Data.count_expert_doctor_fact;
                Dgv.Rows[0].Cells[11].Value = reportCadre.Data.count_expert_doctor_vacancy;
                Dgv.Rows[0].Cells[12].Value = reportCadre.Data.count_grf15;
                Dgv.Rows[0].Cells[13].Value = reportCadre.Data.count_grf16;
                Dgv.Rows[0].Cells[14].Value = reportCadre.Data.count_grf17;
                Dgv.Rows[0].Cells[15].Value = reportCadre.Data.count_grf18;
                Dgv.Rows[0].Cells[16].Value = reportCadre.Data.count_grf19;
                Dgv.Rows[0].Cells[17].Value = reportCadre.Data.count_grf20;
                Dgv.Rows[0].Cells[18].Value = reportCadre.Data.count_grf21;
                Dgv.Rows[0].Cells[19].Value = reportCadre.Data.count_grf22;
                Dgv.Rows[0].Cells[20].Value = reportCadre.Data.count_grf23;
                Dgv.Rows[0].Cells[21].Value = reportCadre.Data.count_grf24;
                Dgv.Rows[0].Cells[22].Value = reportCadre.Data.count_grf25;
                Dgv.Rows[0].Cells[23].Value = reportCadre.Data.count_grf26;
                Dgv.Rows[0].Cells[24].Value = reportCadre.Data.count_specialist_state;
                Dgv.Rows[0].Cells[25].Value = reportCadre.Data.count_specialist_fact;
                Dgv.Rows[0].Cells[26].Value = reportCadre.Data.count_specialist_vacancy;

            }

            
            SetFormula();


        }
        public override void SaveReportDataSourceExcel()
        { }
        public override void SaveReportDataSourceHandle()
        { }

        public void SetFormula()
        {

            try
            {
                Dgv.Rows[0].Cells[2].Value = Math.Round(GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[5].Value) + GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[8].Value) +
                                                        GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[11].Value), 2);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            try
            {
                Dgv.Rows[0].Cells[1].Value = Math.Round(GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[4].Value) + GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[7].Value) +
                                                        GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[10].Value), 2);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            try
            {
                Dgv.Rows[0].Cells[0].Value = Math.Round(GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[3].Value) + GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[6].Value) +
                                                        GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[9].Value), 2);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            try
            {
                Dgv.Rows[0].Cells[5].Value = Math.Round(GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[3].Value) - GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[4].Value), 2);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            
            try
            {
                Dgv.Rows[0].Cells[8].Value = Math.Round(GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[6].Value) - GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[7].Value), 2);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


            try
            {
                Dgv.Rows[0].Cells[11].Value = Math.Round(GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[9].Value) - GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[10].Value), 2);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            
            
            try
            {

                Dgv.Rows[0].Cells[26].Value = Math.Round(GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[24].Value) - GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[25].Value), 2);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


        }


        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource)
        {

        }
        public override void InitReport()
        {
            Report = new ReportCadre { ReportDataList = new ReportCadreDto[ThemesList.Count], IdType = IdReportType };
            int i = 0;
            foreach (var theme in ThemesList.Select(x => x.Key))
            {
                Report.ReportDataList[i++] = new ReportCadreDto { Theme = theme };
                //Console.WriteLine(FilialName);
            }
            SetFormula();
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
            var inReport = report as ReportCadre;

            var index = Report.ReportDataList.ToList().FindIndex(x => x.Theme == Cmb.Text);
            var inTheme = inReport.ReportDataList.Single(x => x.Theme == Cmb.Text);
            Report.ReportDataList[index] = inTheme;

        }
        public override void SaveToDb()
        {
            SetFormula();
            var request = new SaveReportRequest
            {
                Body = new SaveReportRequestBody
                {
                    filialCode = CurrentUser.FilialCode,
                    idUser = CurrentUser.IdUser,
                    report = Report,
                    yymm = Report.Yymm,
                    reportType = ReportType.Cadre
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as ReportCadre;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
        }
        public override void ToExcel(string filename, string filialName)
        {

                var excel = new ExcelCadreCreator(filename, ExcelForm.cadre, Report.Yymm, filialName, Client, FilialCode);
                excel.CreateReport(Report, null);
        }
        public override string ValidReport() { return ""; }
        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            Dgv.AllowUserToAddRows = false;
            Dgv.ColumnHeadersVisible = true;

            Dgv.Columns.Clear();
            Dgv.Rows.Clear();

            List<string> columns = null;
            if (form == "Отдел ЗПЗ и ЭКМП")
            {
                columns = zpz_ekmp;
            }
            else if (form == "ОИ и ЗПЗ")
            {
                columns = oi_zpz;

            }

            foreach (var clmn in columns)
            {
                var column = new DataGridViewTextBoxColumn
                {
                    HeaderText = clmn,
                    DataPropertyName = "Indicator",
                    Name = "Indicator",
                    SortMode = DataGridViewColumnSortMode.NotSortable,
                    DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Azure }
                };

                Dgv.Columns.Add(column);
            }

            Dgv.Rows.Add();

            
            // красим диапазоны колонок в соответствии с шаблоном Excel
            for (int i = 0; i < 27; i++)
            {
                //Dgv.Columns[i].Width = ;
                Dgv.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            for (int j = 0; j < 2; j++)
            {
                Dgv.Columns[j].DefaultCellStyle.BackColor = Color.FromArgb(253, 233, 217);
            }
            for (int j = 3; j < 5; j++)
            {
                Dgv.Columns[j].DefaultCellStyle.BackColor = Color.FromArgb(242, 220, 219);
            }
            for (int j = 6; j < 8; j++)
            {
                Dgv.Columns[j].DefaultCellStyle.BackColor = Color.FromArgb(216, 228, 188);
            }
            for (int j = 9; j < 11; j++)
            {
                Dgv.Columns[j].DefaultCellStyle.BackColor = Color.FromArgb(197, 217, 241);
            }
            for (int j = 12; j < 18; j++)
            {
                Dgv.Columns[j].DefaultCellStyle.BackColor = Color.FromArgb(218, 238, 243);
            }
            for (int j = 18; j < 24; j++)
            {
                Dgv.Columns[j].DefaultCellStyle.BackColor = Color.FromArgb(228, 223, 236);
            }
            for (int j = 24; j < 26; j++)
            {
                Dgv.Columns[j].DefaultCellStyle.BackColor = Color.FromArgb(242, 242, 242);
            }
            // конец покраски

            Dgv.Columns[0].ReadOnly =
            Dgv.Columns[1].ReadOnly =
            Dgv.Columns[2].ReadOnly =
            Dgv.Columns[5].ReadOnly =
            Dgv.Columns[8].ReadOnly =
            Dgv.Columns[11].ReadOnly = 
            Dgv.Columns[26].ReadOnly = true;


            //Dgv.Columns[1].DefaultCellStyle.BackColor =
            //Dgv.Columns[2].DefaultCellStyle.BackColor =
            Dgv.Columns[2].DefaultCellStyle.BackColor = Color.FromArgb(252, 213, 180);
            Dgv.Columns[5].DefaultCellStyle.BackColor = Color.FromArgb(230, 184, 183);
            Dgv.Columns[8].DefaultCellStyle.BackColor = Color.FromArgb(196, 215, 155);
            Dgv.Columns[11].DefaultCellStyle.BackColor = Color.FromArgb(141, 180, 226);
            Dgv.Columns[26].DefaultCellStyle.BackColor = Color.FromArgb(217, 217, 217);
            Dgv.AutoSize = true;
            Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

        }
        protected override void FillReport(string form)
        {
            var reportCadre = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            var row = Dgv.Rows[0];
            reportCadre.Data = new ReportCadreDataDto
            {
                count_itog_state = GlobalUtils.TryParseDecimal(row.Cells[0].Value),
                count_itog_fact = GlobalUtils.TryParseDecimal(row.Cells[1].Value),
                count_itog_vacancy = GlobalUtils.TryParseDecimal(row.Cells[2].Value),
                count_leader_state = GlobalUtils.TryParseDecimal(row.Cells[3].Value),
                count_leader_fact = GlobalUtils.TryParseDecimal(row.Cells[4].Value),
                count_leader_vacancy = GlobalUtils.TryParseDecimal(row.Cells[5].Value),
                count_deputy_leader_state = GlobalUtils.TryParseDecimal(row.Cells[6].Value),
                count_deputy_leader_fact = GlobalUtils.TryParseDecimal(row.Cells[7].Value),
                count_deputy_leader_vacancy = GlobalUtils.TryParseDecimal(row.Cells[8].Value),
                count_expert_doctor_state = GlobalUtils.TryParseDecimal(row.Cells[9].Value),
                count_expert_doctor_fact = GlobalUtils.TryParseDecimal(row.Cells[10].Value),
                count_expert_doctor_vacancy = GlobalUtils.TryParseDecimal(row.Cells[11].Value),
                count_grf15 = GlobalUtils.TryParseInt(row.Cells[12].Value),
                count_grf16 = GlobalUtils.TryParseInt(row.Cells[13].Value),
                count_grf17 = GlobalUtils.TryParseInt(row.Cells[14].Value),
                count_grf18 = GlobalUtils.TryParseInt(row.Cells[15].Value),
                count_grf19 = GlobalUtils.TryParseInt(row.Cells[16].Value),
                count_grf20 = GlobalUtils.TryParseInt(row.Cells[17].Value),
                count_grf21 = GlobalUtils.TryParseInt(row.Cells[18].Value),
                count_grf22 = GlobalUtils.TryParseInt(row.Cells[19].Value),
                count_grf23 = GlobalUtils.TryParseInt(row.Cells[20].Value),
                count_grf24 = GlobalUtils.TryParseInt(row.Cells[21].Value),
                count_grf25 = GlobalUtils.TryParseInt(row.Cells[22].Value),
                count_grf26 = GlobalUtils.TryParseInt(row.Cells[23].Value),
                count_specialist_state = GlobalUtils.TryParseDecimal(row.Cells[24].Value),
                count_specialist_fact = GlobalUtils.TryParseDecimal(row.Cells[25].Value),
                count_specialist_vacancy = GlobalUtils.TryParseDecimal(row.Cells[26].Value),
            };
        }
    }
}
