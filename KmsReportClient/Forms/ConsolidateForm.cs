using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using KmsReportClient.Excel.Creator.Consolidate;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Model.Enums;
using KmsReportClient.Support;
using NLog.Fluent;

namespace KmsReportClient.Forms
{
    public partial class ConsolidateForm : Form
    {
        private const string SummaryFilialName = "ООО «Капитал МС»";
        private const string SummaryFilialCode = "RU";
        private static readonly Dictionary<string, string> region_name = new Dictionary<string, string>
             {
                { "RU-AL", "Горно-Алтайск" },
                { "RU-ALT", "Алтай" },
                { "RU-ARK", "Архангельск" },
                { "RU-BA", "Башкортостан" },
                { "RU-BU", "Бурятия" },
                { "RU-KB", "КБР" },
                { "RU-KDA", "Краснодар" },
                { "RU-KGD", "Калининград" },
                { "RU-KGN", "Курган" },
                { "RU-KHA", "Хабаровск" },
                { "RU-KHM", "ХМАО" },
                { "RU-KIR", "Киров" },
                { "RU-KO", "Коми" },
                { "RU-KOS", "Кострома" },
                { "RU-LEN", "Ленобласть" },
                { "RU-LIP", "Липецк" },
                { "RU-MO", "Мордовия" },
                { "RU-MOS", "Мособласть" },
                { "RU-MOW", "Москва" },
                { "RU-NEN", "НАО" },
                { "RU-NIZ", "Н.Новгород" },
                { "RU-OMS", "Омск" },
                { "RU-ORE", "Оренбург" },
                { "RU-PER", "Пермь" },
                { "RU-PNZ", "Пенза" },
                { "RU-ROS", "Ростов" },
                { "RU-RYA", "Рязань" },
                { "RU-SA", "Якутия" },
                { "RU-SAR", "Саратов" },
                { "RU-SE", "РСО-Алания" },
                { "RU-SMO", "Смоленск" },
                { "RU-SPE", "Санкт-Петербург" },
                { "RU-TUL", "Тула" },
                { "RU-TVE", "Тверь" },
                { "RU-TY", "Тыва" },
                { "RU-TYU", "Тюмень" },
                { "RU-UD", "Удмуртия" },
                { "RU-ULY", "Ульяновск" },
                { "RU-VGG", "Волгоград" },
                { "RU-VLA", "Владимир" },
                { "RU-YAR", "Ярославль" },
                { "RU-YEV", "ЕАО" },
             };

        private static readonly ConsolidateReport[] FolderReports = { ConsolidateReport.ZpzWebSite, 
                                                                      ConsolidateReport.ZpzWebSite2023,
                                                                      ConsolidateReport.ZpzWebSite2025,
                                                                      ConsolidateReport.ViolationsOfAppeals,
                                                                      ConsolidateReport.FFOMSTargetedExp,
                                                                      ConsolidateReport.FFOMSPersonnel,
                                                                      ConsolidateReport.FFOMSViolMEE,
                                                                      ConsolidateReport.FFOMSViolEKMP,
                                                                      ConsolidateReport.FFOMSVerifyPlan,
                                                                      ConsolidateReport.FFOMSMonthlyVol,
                                                                      

                                                                    };

        private readonly EndpointSoapClient _client;
        private readonly string _filialName;
        private readonly List<KmsReportDictionary> _regions;
        private readonly ConsolidateReport _report;

        public ConsolidateForm(EndpointSoapClient client, List<KmsReportDictionary> regions,
            ConsolidateReport report, string filialName)
        {
            InitializeComponent();
            this._client = client;
            this._report = report;
            this._regions = regions;
            this._filialName = filialName;
        }

        private void ConsolidateForm_Load(object sender, EventArgs e)
        {
            int year = DateTime.Today.Year;
            nudStart.Value = year;
            nudEnd.Value = year;
            nudSingle.Value = year;

            var monthsStart = GlobalConst.Months;
            var monthsEnd = GlobalConst.Months;

            cmbStart.DataSource = monthsStart;
            cmbEnd.DataSource = GlobalConst.Months.Clone();

            cmbRegion.DisplayMember = "Value";
            cmbRegion.ValueMember = "Key";

            if (_report == ConsolidateReport.ConsolidateOped)
            {
                _regions.Add(new KmsReportDictionary { Key = "Все филиалы", Value = "Все филиалы" });

            }
            cmbRegion.DataSource = _regions;

            switch (_report)
            {
                case ConsolidateReport.ConsolidateOpedUnplanned:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать свод по отчету Внеплановый ОПЭД";
                    saveFileDialog1.FileName = "Свод по отчету Внеплановый ОПЭД";
                    break;
                case ConsolidateReport.ConsolidateCadreT1:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать свод по отчету Кадры Отдел ЗПЗ и ЭКМП";
                    saveFileDialog1.FileName = "Свод по отчету Кадры Отдел ЗПЗ и ЭКМП";
                    break;
                case ConsolidateReport.ConsolidateCadreT2:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать свод по отчету Кадры ОИ и ЗПЗ";
                    saveFileDialog1.FileName = "Свод по отчету Кадры ОИ и ЗПЗ";
                    break;
                case ConsolidateReport.Consolidate262T1:
                    labelStart.Text = "Год";
                    panelSt.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать свод к табл.1 по отчетной форме к приказу 262";
                    saveFileDialog1.FileName = "Свод к табл.1 по форме 262";
                    break;
                case ConsolidateReport.Consolidate262T2:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать свод к табл.2 по отчетной форме к приказу 262";
                    saveFileDialog1.FileName = "Свод к табл.2 по форме 262";
                    break;
                case ConsolidateReport.Consolidate262T3:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать свод к табл.3 по отчетной форме к приказу 262";
                    saveFileDialog1.FileName = "Свод к табл.3 по форме 262";
                    break;
                case ConsolidateReport.ConsolidateFilial294:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = true;
                    btnDo.Text = "Сформировать сводную таблицу по форме 294";
                    saveFileDialog1.FileName = "Свод по форме 294";
                    break;
                case ConsolidateReport.ConsolidateFull294:
                    labelStart.Text = "Год";
                    panelSt.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать итоговый отчет по филиалам по форме 294";
                    saveFileDialog1.FileName = "Итоговый отчет по форме 294";
                    break;
                case ConsolidateReport.ZpzWebSite:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать отчет ЗПЗ для сайта";
                    saveFileDialog1.FileName = "Сводный отчет ЗПЗ для сайта";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;
                case ConsolidateReport.ZpzWebSite2023:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать отчет ЗПЗ для сайта";
                    saveFileDialog1.FileName = "Сводный отчет ЗПЗ для сайта";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;
                case ConsolidateReport.FFOMSTargetedExp:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать отчет по Внеплановым экспертизам";
                    saveFileDialog1.FileName = "Отчет по Внеплановым экспертизам";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;
                case ConsolidateReport.FFOMSOncoCT:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать сводный отчет Онкология ХТ";
                    saveFileDialog1.FileName = "Сводный отчет Онкология ХТ";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;
                case ConsolidateReport.FFOMSPersonnel:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать отчет Кадры";
                    saveFileDialog1.FileName = "Отчет Кадры";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;
                case ConsolidateReport.FFOMSLethalEKMP:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать отчет Летальные ЭКМП";
                    saveFileDialog1.FileName = "Отчет Летальные ЭКМП";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;
                case ConsolidateReport.ZpzTable5:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать отчет Результаты МЭК пофилиально";
                    saveFileDialog1.FileName = "Отчет Результаты МЭК пофилиально";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;
                case ConsolidateReport.Zpz10Cons:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать свод по таблице 10 формы ЗПЗ";
                    saveFileDialog1.FileName = "Свод по таблице 10 формы ЗПЗ 118н";
                    cmbStart.DataSource = GlobalConst.Months;
                    break;
                case ConsolidateReport.Zpz10FilialCons:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать свод пофилиально по таблице 10 формы ЗПЗ";
                    saveFileDialog1.FileName = "Пофилиальный свод по таблице 10 формы ЗПЗ 118н";
                    cmbStart.DataSource = GlobalConst.Months;
                    break;
                case ConsolidateReport.Zpz10FilialGrowCons:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать свод по филиалам наростом по таблице 10 формы ЗПЗ";
                    saveFileDialog1.FileName = "Свод по филиалам наростом по таблице 10 формы ЗПЗ 118н";
                    cmbStart.DataSource = GlobalConst.Months;
                    break;
                case ConsolidateReport.FFOMSVolumesByTypes:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать отчет Объемы по видам помощи";
                    saveFileDialog1.FileName = "Отчет Объемы по видам помощи";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;
                case ConsolidateReport.ZpzWebSite2025:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать отчет ЗПЗ 118н для сайта";
                    saveFileDialog1.FileName = "Сводный отчет ЗПЗ 118н для сайта";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;
                case ConsolidateReport.ViolationsOfAppeals:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать отчет Нарушения по обращениям ЗЛ";
                    saveFileDialog1.FileName = "Нарушения по обращениям ЗЛ";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;
                case ConsolidateReport.FFOMSViolMEE:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать отчет Нарушения МЭЭ";
                    saveFileDialog1.FileName = "Нарушения МЭЭ";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;
                case ConsolidateReport.FFOMSViolEKMP:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать отчет Нарушения ЭКМП";
                    saveFileDialog1.FileName = "Нарушения ЭКМП";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;
                case ConsolidateReport.FFOMSVerifyPlan:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать отчет Планы проверок";
                    saveFileDialog1.FileName = "Планы проверок";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;
                case ConsolidateReport.FFOMSMonthlyVol:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать отчет Объемы ежемесячные";
                    saveFileDialog1.FileName = "Объемы ежемесячные";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;
                case ConsolidateReport.ControlZpzMonthly:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать сводный отчет для контроля ЗПЗ";
                    saveFileDialog1.FileName = "Сводный отчет для контроля ЗПЗ";
                    break;
                case ConsolidateReport.ControlZpzQuarterly:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать сводный отчет для контроля ЗПЗ (квартальный)";
                    saveFileDialog1.FileName = "Сводный отчет для контроля ЗПЗ (квартальный)";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;
                case ConsolidateReport.ControlZpz2023Quarterly:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать сводный отчет для контроля ЗПЗ 2023(квартальный)";
                    saveFileDialog1.FileName = "Сводный отчет для контроля ЗПЗ 2023(квартальный)";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;
                case ConsolidateReport.ControlZpz2023FullQuarterly:
                    labelStart.Text = "Год";
                    panelSt.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать отчет для контроля ЗПЗ 2023(за весь год)";
                    saveFileDialog1.FileName = "Сводный отчет для контроля ЗПЗ 2023(за весь год)";
                    break;
                case ConsolidateReport.ControlZpz2023SingleQuarterly:
                    labelStart.Text = "Год";
                    panelSt.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Проверочная таблица ЗПЗ 2023(за весь год)";
                    saveFileDialog1.FileName = "Проверочная таблица ЗПЗ 2023(за весь год)";
                    break;
                case ConsolidateReport.ControlZpz2025Quarterly:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать сводный отчет для контроля ЗПЗ 118н(квартальный)";
                    saveFileDialog1.FileName = "Сводный отчет для контроля ЗПЗ 118н(квартальный)";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;
                case ConsolidateReport.ControlZpz2025FullQuarterly:
                    labelStart.Text = "Год";
                    panelSt.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать отчет для контроля ЗПЗ 118н(за весь год)";
                    saveFileDialog1.FileName = "Сводный отчет для контроля ЗПЗ 118н(за весь год)";
                    break;
                case ConsolidateReport.ControlZpz2025SingleQuarterly:
                    labelStart.Text = "Год";
                    panelSt.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Проверочная таблица ЗПЗ 118н(за весь год)";
                    saveFileDialog1.FileName = "Проверочная таблица ЗПЗ 118н(за весь год)";
                    break;
                case ConsolidateReport.Onko:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать сводный отчет по онкологии";
                    saveFileDialog1.FileName = "Сводный отчет по онкологии";
                    break;
                case ConsolidateReport.OnkoQuarterly:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать сводный отчет по онкологии (квартальный)";
                    saveFileDialog1.FileName = "Сводный отчет по онкологии (квартальный)";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;

                case ConsolidateReport.Cardio:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать сводный отчет по сердечно-сосудистым заболеваниям (квартальный)";
                    saveFileDialog1.FileName = "Сводный отчет по сердечно-сосудистым заболеваниям (квартальный)";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;

                case ConsolidateReport.Disp:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать сводный отчет Диспансеризация(квартальный)";
                    saveFileDialog1.FileName = "Сводный отчет Диспансеризация (квартальный)";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;

                case ConsolidateReport.Letal:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать сводный отчет Летальные";
                    saveFileDialog1.FileName = "Сводный отчет Летальные";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;

                case ConsolidateReport.Letal2023:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать сводный отчет Летальные";
                    saveFileDialog1.FileName = "Сводный отчет Летальные";
                    cmbStart.DataSource = GlobalConst.Periods;
                    break;

                case ConsolidateReport.CnpnQuarterly:
                    labelStart.Text = "Период";
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    nudSingle.Visible = false;
                    cmbStart.DataSource = GlobalConst.Periods;
                    btnDo.Text = "Сформировать сводный отчет  об исполнении ЦПНП";
                    saveFileDialog1.FileName = "Cводный отчет  об исполнении ЦПНП";
                    break;

                case ConsolidateReport.ConsolidateOped:
                    //oped
                    labelStart.Text = "Период начала";
                    panelEnd.Visible = true;
                    panelRegion.Visible = true;
                    nudSingle.Visible = false;
                    cmbStart.DataSource = GlobalConst.Months;
                    btnDo.Text = "Сформировать сводный отчёт ОПЭД";
                    saveFileDialog1.FileName = "Сводный отчёт ОПЭД";
                    break;
                case ConsolidateReport.CnpnMonthly:
                    labelStart.Text = "Период";
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    nudSingle.Visible = false;
                    cmbStart.DataSource = GlobalConst.Months;
                    btnDo.Text = "Сформировать сводный отчет об исполнении ЦПНП";
                    saveFileDialog1.FileName = "Cводный отчет об исполнении ЦПНП";
                    break;

                case ConsolidateReport.ConsolidateVSS:
                    labelStart.Text = "Период";
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    nudSingle.Visible = false;
                    cmbStart.DataSource = GlobalConst.Months;
                    btnDo.Text = "Сформировать сводный отчёт Мониторинг ВСС";
                    saveFileDialog1.FileName = "Cводный отчет Мониторинг ВСС";
                    break;

                case ConsolidateReport.ConsolidateVCR:
                    labelStart.Text = "Период";
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    nudSingle.Visible = false;
                    cmbStart.DataSource = GlobalConst.Periods;
                    btnDo.Text = "Сформировать сводный отчёт Мониторинг ВСС 2024";
                    saveFileDialog1.FileName = "Cводный отчет Мониторинг ВСС 2024";
                    break;

                case ConsolidateReport.ConsolidateOpedQ:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать сводный отчет ОПЭД внеплановый поквартально";
                    saveFileDialog1.FileName = "Сводный отчет ОПЭД внеплановый";
                    cmbStart.DataSource = GlobalConst.PeriodsQ;
                    break;

                case ConsolidateReport.ConsolidateCPNP2Q:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать сводный отчет об исполнении ЦПНП";
                    saveFileDialog1.FileName = "Сводный отчет об исполнении ЦПНП";
                    cmbStart.DataSource = GlobalConst.PeriodsQ;
                    break;

                case ConsolidateReport.ConsOpedFinance1:
                    labelStart.Text = "Период";
                    //nudSingle.Visible = false;
                    nudEnd.Visible = false;
                    cmbEnd.Visible = false;
                    cmbStart.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    nudStart.Visible = false;
                    btnDo.Text = "Сформировать сводный отчёт ОПЭД финансы 1";
                    saveFileDialog1.FileName = "Сводный отчёт ОПЭД финансы 1";
                    cmbStart.DataSource = GlobalConst.PeriodsQ;
                    labelStart.Text = "Год";
                    break;

                case ConsolidateReport.ConsOpedFinance2:
                    labelStart.Text = "Период";
                    //nudSingle.Visible = false;
                    nudEnd.Visible = false;
                    cmbEnd.Visible = false;
                    cmbStart.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    nudStart.Visible = false;
                    btnDo.Text = "Сформировать сводный отчёт ОПЭД финансы 2";
                    saveFileDialog1.FileName = "Сводный отчёт ОПЭД финансы 2";
                    cmbStart.DataSource = GlobalConst.PeriodsQ;
                    labelStart.Text = "Год";
                    break;

                case ConsolidateReport.ConsOpedFinance3:
                    labelStart.Text = "Период";
                    //nudSingle.Visible = false;
                    nudEnd.Visible = false;
                    cmbEnd.Visible = false;
                    cmbStart.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    nudStart.Visible = false;
                    btnDo.Text = "Сформировать сводный отчёт ОПЭД финансы 3";
                    saveFileDialog1.FileName = "Сводный отчёт ОПЭД финансы 3";
                    cmbStart.DataSource = GlobalConst.PeriodsQ;
                    labelStart.Text = "Год";
                    break;

                case ConsolidateReport.ConsQuantityFP:
                    labelStart.Text = "Период";
                    //nudSingle.Visible = false;
                    nudEnd.Visible = false;
                    cmbEnd.Visible = false;
                    cmbStart.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    nudStart.Visible = false;
                    btnDo.Text = "Сформировать сводный отчет по выполнению плана";
                    saveFileDialog1.FileName = "Выполнение плана по численности";
                    cmbStart.DataSource = GlobalConst.PeriodsQ;
                    labelStart.Text = "Год";
                    break;


                case ConsolidateReport.ConsQuantityAR:
                    labelStart.Text = "Период";
                    //nudSingle.Visible = false;
                    nudEnd.Visible = false;
                    cmbEnd.Visible = false;
                    cmbStart.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    nudStart.Visible = false;
                    btnDo.Text = "Сформировать сводный отчёт по численности ВЗ и УД";
                    saveFileDialog1.FileName = "Сводный отчёт Численность по ВЗ и УД";
                    cmbStart.DataSource = GlobalConst.PeriodsQ;
                    labelStart.Text = "Год";
                    break;

                case ConsolidateReport.ConsQuantityInformation:
                    labelStart.Text = "Период";
                    //nudSingle.Visible = false;
                    nudEnd.Visible = false;
                    cmbEnd.Visible = false;
                    cmbStart.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    nudStart.Visible = false;
                    btnDo.Text = "Сформировать свод сведения по численности";
                    saveFileDialog1.FileName = "Свод сведения о численности";
                    cmbStart.DataSource = GlobalConst.PeriodsQ;
                    labelStart.Text = "Год";
                    break;

                case ConsolidateReport.ConsPropsal:
                    labelStart.Text = "Период";
                    nudSingle.Visible = false;
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    btnDo.Text = "Сформировать сводный отчет о предложениях";
                    saveFileDialog1.FileName = "Сводный отчет о предложениях";
                    cmbStart.DataSource = GlobalConst.PeriodsQ;
                    break;

                case ConsolidateReport.ConsolidateVCRFilial:
                    labelStart.Text = "Период";
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    nudSingle.Visible = false;
                    cmbStart.DataSource = GlobalConst.Periods;
                    btnDo.Text = "Сформировать сводный отчет Мониторинг ВСС 2024 пофилиально";
                    saveFileDialog1.FileName = "Пофилиальный Мониторинг ВСС 2024";
                    break;

                case ConsolidateReport.ConsQuantityFilial:
                    labelStart.Text = "Период";
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    nudSingle.Visible = false;

                    btnDo.Text = "Сформировать сводный отчет Численность пофилиально";
                    saveFileDialog1.FileName = "Численность по всем филиалам";
                    break;

                case ConsolidateReport.ConsQuantityQ:
                    labelStart.Text = "Период";
                    panelEnd.Visible = false;
                    panelRegion.Visible = false;
                    nudSingle.Visible = false;
                    cmbStart.DataSource = GlobalConst.Periods;
                    btnDo.Text = "Сформировать сводный отчёт численность за период";
                    saveFileDialog1.FileName = "Квартальный сводный отчет численность";
                    break;
            }
        }

        private int GetYymm(string monthForm, int yearForm)
        {
            int month = Array.IndexOf(GlobalConst.Months, monthForm) + 1;
            return (yearForm - 2000) * 100 + month;
        }

        private void BtnDo_Click(object sender, EventArgs e)
        {
            var dialogResult = FolderReports.Contains(_report) ?
                folderBrowserDialog1.ShowDialog() :
                saveFileDialog1.ShowDialog();

            if (dialogResult == DialogResult.OK)
            {
                var waitingForm = new WaitingForm();
                waitingForm.Show();
                switch (_report)
                {
                    case ConsolidateReport.ConsQuantityFP:
                        CreateConsolidateQuantityFactPlan();
                        break;
                    case ConsolidateReport.ConsQuantityInformation:
                        CreateConsolidateQuantityInfo();
                        break;
                    case ConsolidateReport.ConsolidateVCRFilial:
                        CreateReportVCRFilial();
                        break;
                    case ConsolidateReport.ConsolidateCadreT1:
                        CreateReportCadreT1();
                        break;
                    case ConsolidateReport.ConsolidateCadreT2:
                        CreateReportCadreT2();
                        break;
                    case ConsolidateReport.ConsolidateOpedUnplanned:
                        CreateReportOpedUnplanned();
                        break;
                    case ConsolidateReport.Consolidate262T1:
                        CreateReport262T1();
                        break;
                    case ConsolidateReport.Consolidate262T2:
                        CreateReport262T2();
                        break;
                    case ConsolidateReport.Consolidate262T3:
                        CreateReport262T3();
                        break;
                    case ConsolidateReport.ConsolidateFilial294:
                        CreateFilial294();
                        break;
                    case ConsolidateReport.ConsolidateFull294:
                        CreateFull294();
                        break;
                    case ConsolidateReport.ControlZpzMonthly:
                        CreateControlZpz(true);
                        break;
                    case ConsolidateReport.ControlZpzQuarterly:
                        CreateControlZpz(false);
                        break;
                    case ConsolidateReport.ControlZpz2023Quarterly:
                        CreateControlZpz2023(false);
                        break;
                    case ConsolidateReport.ControlZpz2023FullQuarterly:
                        CreateControlZpz2023Full();
                        break;
                    case ConsolidateReport.ControlZpz2025Quarterly:
                        CreateControlZpz2025(false);
                        break;
                    case ConsolidateReport.ControlZpz2025FullQuarterly:
                        CreateControlZpz2025Full();
                        break;
                    case ConsolidateReport.ConsQuantityFilial:
                        CreateReportConsQuantityFilial();
                        break;
                    case ConsolidateReport.ConsQuantityAR:
                        CreateConsolidateQuantityAddRemove();
                        break;
                    case ConsolidateReport.FFOMSVolumesByTypes:
                        CreateFFOMSVolumesByTypes();
                        break;
                    case ConsolidateReport.FFOMSLethalEKMP:
                        CreateFFOMSLethalEKMP();
                        break;
                    case ConsolidateReport.ZpzTable5:
                        CreateZpzTable5();
                        break;
                    case ConsolidateReport.ConsQuantityQ:
                        CreateConsolidateQuantityQ();
                        break;
                    case ConsolidateReport.ControlZpz2023SingleQuarterly:
                        CreateControlZpz2023Single();
                        break;
                    case ConsolidateReport.ControlZpz2025SingleQuarterly:
                        CreateControlZpz2025Single();
                        break;
                    case ConsolidateReport.ZpzWebSite:
                        CreateZpzWebSite();
                        break;
                    case ConsolidateReport.ZpzWebSite2023:
                        CreateZpzWebSite2023();
                        break;
                    case ConsolidateReport.FFOMSTargetedExp:
                        CreateFFOMSTargetedExp();
                        break;
                    case ConsolidateReport.FFOMSOncoCT:
                        CreateFFOMSOncoCT();
                        break;
                    case ConsolidateReport.FFOMSPersonnel:
                        CreateFFOMSPersonnel();
                        break;
                    case ConsolidateReport.ViolationsOfAppeals:
                        CreateViolationsOfAppeals();
                        break;
                    case ConsolidateReport.FFOMSViolMEE:
                        CreateFFOMSViolMEE();
                        break;
                    case ConsolidateReport.FFOMSViolEKMP:
                        CreateFFOMSViolEKMP();
                        break;
                    case ConsolidateReport.FFOMSVerifyPlan:
                        CreateFFOMSVerifyPlan();
                        break;
                    case ConsolidateReport.FFOMSMonthlyVol:
                        CreateFFOMSMonthlyVol();
                        break;
                    case ConsolidateReport.ZpzWebSite2025:
                        CreateZpzWebSite2025();
                        break;
                    case ConsolidateReport.Onko:
                        CreateOnko(true);
                        break;
                    case ConsolidateReport.OnkoQuarterly:
                        CreateOnko(false);
                        break;
                    case ConsolidateReport.CnpnQuarterly:
                        CreateCReportCpnp();
                        break;

                    case ConsolidateReport.CnpnMonthly:
                        CreateCReportCpnpMonth();
                        break;
                    case ConsolidateReport.Cardio:
                        CreateCReportCardio();
                        break;
                    case ConsolidateReport.Disp:
                        CreateCReportDisp();
                        break;
                    case ConsolidateReport.Letal:
                        CreateCReportLetal();
                        break;
                    case ConsolidateReport.Letal2023:
                        CreateCReportLetal();
                        break;
                    case ConsolidateReport.ConsolidateOped:
                        CreateCOped();
                        break;
                    case ConsolidateReport.ConsolidateVSS:
                        CreateCVSS();
                        break;
                    case ConsolidateReport.ConsolidateVCR:
                        CreateCVCR();
                        break;
                    case ConsolidateReport.ConsolidateOpedQ:
                        CreateCOpedQ();
                        break;
                    case ConsolidateReport.ConsolidateCPNP2Q:
                        CreateCPNP2Q();
                        break;
                    case ConsolidateReport.ConsOpedFinance1:
                        ConsolidateOpedFinance1();
                        break;
                    case ConsolidateReport.ConsOpedFinance2:
                        ConsolidateOpedFinance2();
                        break;
                    case ConsolidateReport.ConsOpedFinance3:
                        ConsolidateOpedFinance3();
                        break;
                    case ConsolidateReport.ConsPropsal:
                        ConsolidateProposal();
                        break;
                    case ConsolidateReport.Zpz10Cons:
                        CreateZpz10Cons();
                        break;
                    case ConsolidateReport.Zpz10FilialCons:
                        CreateZpz10FilialCons();
                        break;
                    case ConsolidateReport.Zpz10FilialGrowCons:
                        CreateZpz10FilialGrowCons();
                        break;
                }
                waitingForm.Close();
            }

            Close();
        }

        private void ConsolidateProposal()
        {
            string yymm = GetYymmQuarterly2();

            var data = _client.ConsolidateProposalCollect(yymm);

            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //string header = $"За {cmbStart.Text} {nudStart.Value} года";
            var excel = new ExcelConsoldatePropasalCreator(saveFileDialog1.FileName, "", _filialName);

            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }

        private void ConsolidateOpedFinance1()
        {
            string year = nudSingle.Value.ToString();

            var data = _client.ConsolidateOpedFinance1(year);

            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var excel = new ExcelConsolidateOpenFinance1(saveFileDialog1.FileName, "", _filialName);
            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }

        private void ConsolidateOpedFinance2()
        {
            string year = nudSingle.Value.ToString();

            var data = _client.ConsolidateOpedFinance2(year);

            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var excel = new ExcelConsolidateOpenFinance2(saveFileDialog1.FileName, "", _filialName);
            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }

        private void ConsolidateOpedFinance3()
        {
            string year = nudSingle.Value.ToString();


            var data = _client.ConsolidateOpedFinance3(year);

            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var excel = new ExcelConsolidateOpenFinance3(saveFileDialog1.FileName, "", _filialName);
            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }

        private void CreateCPNP2Q()
        {
            string yymm = GetYymmQuarterly2();

            var data = _client.ConsolidateCPNP2QCollect(yymm);

            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string header = $"За {cmbStart.Text} {nudStart.Value} года";
            var excel = new ExcelConsoldateCPNPQ2Creator(saveFileDialog1.FileName, header, _filialName);

            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }

        private void CreateCOpedQ()
        {
            string yymm = GetYymmQuarterly2();

            var data = _client.ConsolidateOpedQCollect(yymm);

            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var excel = new ExcelConsolidateOpedQCreator(saveFileDialog1.FileName, "", _filialName);

            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);

        }

        private void CreateZpz10Cons()
        {
            string yymm = GetYymm(cmbStart.Text, Convert.ToInt32(nudStart.Value)).ToString();

            var data = _client.CreateConsolidateZpzTable10(yymm);

            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var excel = new ExcelConsolidateZpzTable10Creator(saveFileDialog1.FileName, "", _filialName, yymm);

            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);

        }

        private void CreateZpz10FilialCons()
        {
            string yymm = GetYymmQuarterly();

            var data = _client.CreateConsolidateZpzTable10Filial(yymm);

            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var excel = new ExcelConsolidateZpzTable10FilialCreator(saveFileDialog1.FileName, "", _filialName, yymm);

            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);

        }

        private void CreateZpz10FilialGrowCons()
        {
            string yymm = GetYymmQuarterly();

            var data = _client.CreateConsolidateZpzTable10FilialGrow(yymm);

            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var excel = new ExcelConsolidateZpzTable10FilialGrowCreator(saveFileDialog1.FileName, "", _filialName, yymm);

            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);

        }

        private void CreateCVSS()
        {
            string yymm = GetYymm(cmbStart.Text, Convert.ToInt32(nudStart.Value)).ToString();




            string mm = YymmUtils.GetMonth(yymm.Substring(2));

            var data = _client.CreateReportVSS(yymm);


            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var excel = new ExcelConsolidateVSSCreator(saveFileDialog1.FileName, "", _filialName, mm);

            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }

        private void CreateCVCR()
        {
            string yymm = GetYymmQuarterly();




            string mm = YymmUtils.GetMonth(yymm.Substring(2));

            var data = _client.CreateReportVCR(yymm);


            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var excel = new ExcelConsolidateVCRCreator(saveFileDialog1.FileName, "", _filialName, mm);

            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }

        private void CreateZpzWebSite()
        {
            string yymm = GetYymmQuarterly();
            string folder = folderBrowserDialog1.SelectedPath;

            var reports = _client.CreateZpzForWebSite(yymm);

            foreach (var report in reports)
            {
                string filename = folder + $"\\Отчет_для_сайта_{report.Filial}_{yymm}.xlsx";
                string filialName = _regions.Single(x => x.Key == report.Filial).ForeignKey;
                CreateReport(filename, filialName, report);
            }

            var summaryReport = CollectSummaryReport(reports);
            string summaryFilename = folder + $"\\Отчет_для_сайта_суммарный_{yymm}.xlsx";
            CreateReport(summaryFilename, SummaryFilialName, summaryReport);

            GlobalUtils.OpenFileOrDirectory(folder);
        }

        private void CreateZpzWebSite2023()
        {
            string yymm = GetYymmQuarterly();
            string folder = folderBrowserDialog1.SelectedPath;

            var reports = _client.CreateZpzForWebSite2023(yymm);

            foreach (var report in reports)
            {
                string filename = folder + $"\\Отчет_для_сайта_{report.Filial}_{yymm}.xlsx";
                string filialName = _regions.Single(x => x.Key == report.Filial).ForeignKey;
                CreateReport(filename, filialName, report);
            }

            var summaryReport = CollectSummaryReport2023(reports);
            string summaryFilename = folder + $"\\Отчет_для_сайта_суммарный_{yymm}.xlsx";
            CreateReport(summaryFilename, SummaryFilialName, summaryReport);

            GlobalUtils.OpenFileOrDirectory(folder);
        }



         

        private void CreateViolationsOfAppeals()
        {
            string yymm = GetYymmQuarterly();
            string folder = folderBrowserDialog1.SelectedPath;

            var reports = _client.CreateViolationsOfAppeals(yymm);

            foreach (var report in reports)
            {
                string filename = folder + $"\\Нарушения_по_обращениям_ЗЛ_{report.Filial}_{yymm}.xlsx";
                string filialName = _regions.Single(x => x.Key == report.Filial).ForeignKey;
                CreateReport(filename, filialName, report);
            }

            var summaryReport = CollectSummaryViolationsOfAppeals(reports);
            string summaryFilename = folder + $"\\Нарушения_по_обращениям_ЗЛ_суммарный_{yymm}.xlsx";
            CreateReport(summaryFilename, SummaryFilialName, summaryReport);

            GlobalUtils.OpenFileOrDirectory(folder);
        }


        private void CreateFFOMSViolMEE()
        {
            string yymm = GetYymmQuarterly();
            string folder = folderBrowserDialog1.SelectedPath;

            var reports = _client.CreateFFOMSViolMEE(yymm);

            foreach (var report in reports)
            {
                if (region_name.TryGetValue(report.Filial, out string regionName))
                {
                    string filename = folder + $"\\{regionName}_Нарушения МЭЭ {yymm}.xlsx";
                    string filialName = _regions.FirstOrDefault(x => x.Key == report.Filial)?.ForeignKey
                    ?? throw new KeyNotFoundException($"Filial {report.Filial} не найден в _regions");
                    CreateReport(filename, filialName, report);
                }
                else
                {
                    Log.Warn($"Код региона {report.Filial} не найден в словаре region_name");
                }
            }

            var summaryReport = CollectSummaryFFOMSViolMEE(reports);
            string summaryFilename = folder + $"\\Свод_Нарушения МЭЭ {yymm}.xlsx";
            CreateReport(summaryFilename, SummaryFilialName, summaryReport);

            GlobalUtils.OpenFileOrDirectory(folder);
        }


        private void CreateFFOMSViolEKMP()
        {
            string yymm = GetYymmQuarterly();
            string folder = folderBrowserDialog1.SelectedPath;

            var reports = _client.CreateFFOMSViolEKMP(yymm);

            foreach (var report in reports)
            {
                if (region_name.TryGetValue(report.Filial, out string regionName))
                {
                    string filename = folder + $"\\{regionName}_Нарушения ЭКМП {yymm}.xlsx";
                    string filialName = _regions.FirstOrDefault(x => x.Key == report.Filial)?.ForeignKey
                    ?? throw new KeyNotFoundException($"Filial {report.Filial} не найден в _regions");
                    CreateReport(filename, filialName, report);
                }
                else
                {
                    Log.Warn($"Код региона {report.Filial} не найден в словаре region_name");
                }
            }

            var summaryReport = CollectSummaryFFOMSViolEKMP(reports);
            string summaryFilename = folder + $"\\Свод_Нарушения ЭКМП {yymm}.xlsx";
            CreateReport(summaryFilename, SummaryFilialName, summaryReport);

            GlobalUtils.OpenFileOrDirectory(folder);
        }

        private void CreateFFOMSMonthlyVol()
        {
            string yymm = GetYymmQuarterly();
            string folder = folderBrowserDialog1.SelectedPath;

            var reports = _client.CreateFFOMSMonthlyVol(yymm);

            foreach (var report in reports)
            {
                if (region_name.TryGetValue(report.Filial, out string regionName))
                {
                    string filename = folder + $"\\{regionName}_Объемы ежемесячные {yymm}.xlsx";
                    string filialName = _regions.FirstOrDefault(x => x.Key == report.Filial)?.ForeignKey
                    ?? throw new KeyNotFoundException($"Filial {report.Filial} не найден в _regions");
                    CreateReport(filename, filialName, report);
                }
                else
                {
                    Log.Warn($"Код региона {report.Filial} не найден в словаре region_name");
                }
            }

            var summaryReport = CollectSummaryFFOMSMonthlyVol(reports);
            string summaryFilename = folder + $"\\Свод_Объемы ежемесячные {yymm}.xlsx";
            CreateReport(summaryFilename, SummaryFilialName, summaryReport);

            GlobalUtils.OpenFileOrDirectory(folder);
        }


        private void CreateFFOMSVerifyPlan()
        {
            string yymm = GetYymmQuarterly();
            string folder = folderBrowserDialog1.SelectedPath;

            var reports = _client.CreateFFOMSVerifyPlan(yymm);

            foreach (var report in reports)
            {
                if (region_name.TryGetValue(report.Filial, out string regionName))
                {
                    string filename = folder + $"\\{regionName}_Планы проверок {yymm}.xlsx";
                    string filialName = _regions.FirstOrDefault(x => x.Key == report.Filial)?.ForeignKey
                    ?? throw new KeyNotFoundException($"Filial {report.Filial} не найден в _regions");
                    CreateReport(filename, filialName, report);
                }
                else
                {
                    Log.Warn($"Код региона {report.Filial} не найден в словаре region_name");
                }
            }

            var summaryReport = CollectSummaryFFOMSVerifyPlan(reports);
            string summaryFilename = folder + $"\\Свод_Планы проверок {yymm}.xlsx";
            CreateReport(summaryFilename, SummaryFilialName, summaryReport);

            GlobalUtils.OpenFileOrDirectory(folder);
        }


        private void CreateFFOMSTargetedExp()
        {
            string yymm = GetYymmQuarterly();
            string folder = folderBrowserDialog1.SelectedPath;

            var reports = _client.CreateFFOMSTargetedExp(yymm);


            foreach (var report in reports)
            {
                if (region_name.TryGetValue(report.Filial, out string regionName))
                { 
                    string filename = folder + $"\\{regionName}_Внеплановые экспертизы {yymm}.xlsx";
                    string filialName = _regions.FirstOrDefault(x => x.Key == report.Filial)?.ForeignKey
                    ?? throw new KeyNotFoundException($"Filial {report.Filial} не найден в _regions");
                    CreateReport(filename, filialName, report);
                }
                else
                {
                    Log.Warn($"Код региона {report.Filial} не найден в словаре region_name");
                }
            }

            var summaryReport = CollectSummaryFFOMSTargetedExp(reports);
            string summaryFilename = folder + $"\\Свод_Внеплановые экспертизы {yymm}.xlsx";
            CreateReport(summaryFilename, SummaryFilialName, summaryReport);

            GlobalUtils.OpenFileOrDirectory(folder);
        }


        private void CreateFFOMSPersonnel()
        {
            string yymm = GetYymmQuarterly();
            string folder = folderBrowserDialog1.SelectedPath;

            var reports = _client.CreateFFOMSPersonnel(yymm);


            foreach (var report in reports)
            {
                if (region_name.TryGetValue(report.Filial, out string regionName))
                {
                    string filename = folder + $"\\{regionName}_Кадры {yymm}.xlsx";
                    string filialName = _regions.FirstOrDefault(x => x.Key == report.Filial)?.ForeignKey
                    ?? throw new KeyNotFoundException($"Filial {report.Filial} не найден в _regions");
                    CreateReport(filename, filialName, report);
                }
                else
                {
                    Log.Warn($"Код региона {report.Filial} не найден в словаре region_name");
                }
            }

            var summaryReport = CollectSummaryFFOMSPersonnel(reports);
            string summaryFilename = folder + $"\\Свод_Кадры {yymm}.xlsx";
            CreateReport(summaryFilename, SummaryFilialName, summaryReport);

            GlobalUtils.OpenFileOrDirectory(folder);
        }


        private void CreateFFOMSOncoCT()
        {
                string yymm = GetYymmQuarterly();

                var data = _client.CreateFFOMSOncoCT(yymm);
                if (data.Length == 0)
                {
                    MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                foreach (var d in data)
                {
                    d.Filial = _regions.Single(x => x.Key == d.Filial).Value;
                }

                data = data.OrderBy(x => x.Filial).ToArray();

                string filename = saveFileDialog1.FileName;
                var excel = new ExcelFFOMSOncoCTCreator(filename, "", _filialName, yymm);
                excel.CreateReport(data, null);

                GlobalUtils.OpenFileOrDirectory(filename);
            }



        private void CreateControlZpz2025(bool isMonthly)
        {
            string yymm = isMonthly ?
                GetYymm(cmbStart.Text, Convert.ToInt32(nudStart.Value)).ToString() :
                GetYymmQuarterly();

            var data = _client.CreateReportControlZpz2025(yymm, isMonthly);
            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (var d in data)
            {
                d.Filial = _regions.Single(x => x.Key == d.Filial).Value;
            }

            data = data.OrderBy(x => x.Filial).ToArray();

            string filename = saveFileDialog1.FileName;
            var excel = new ExcelControlZpz2025Creator(filename, "", _filialName, yymm);
            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(filename);
        }





        private void CreateZpzWebSite2025()
        {
            string yymm = GetYymmQuarterly();
            string folder = folderBrowserDialog1.SelectedPath;

            var reports = _client.CreateZpzForWebSite2025(yymm);


            foreach (var report in reports)
            {
                string filename = folder + $"\\Отчет_для_сайта_{report.Filial}_{yymm}.xlsx";
                string filialName = _regions.Single(x => x.Key == report.Filial).ForeignKey;
                CreateReport(filename, filialName, report);
            }

            var summaryReport = CollectSummaryReport2025(reports);
            string summaryFilename = folder + $"\\Отчет_для_сайта_суммарный_{yymm}.xlsx";
            CreateReport(summaryFilename, SummaryFilialName, summaryReport);

            GlobalUtils.OpenFileOrDirectory(folder);
        }

        private ZpzForWebSite CollectSummaryReport(ZpzForWebSite[] reports)
        {
            var treatments = reports.SelectMany(x => x.Treatments).GroupBy(x => x.Row).Select(x => new ZpzTreatment
            {
                Row = x.Key,
                Oral = x.Sum(x => x.Oral),
                Written = x.Sum(x => x.Written)
            }).ToArray();
            var complaints = reports.SelectMany(x => x.Complaints).GroupBy(x => x.Row).Select(x => new ZpzTreatment
            {
                Row = x.Key,
                Oral = x.Sum(x => x.Oral),
                Written = x.Sum(x => x.Written)
            }).ToArray();
            var protections = reports.SelectMany(x => x.Protections).GroupBy(x => x.Row).Select(x => new ZpzStatistics
            {
                Row = x.Key,
                Count = x.Sum(x => x.Count)
            }).ToArray();
            var expertises = reports.SelectMany(x => x.Expertises).GroupBy(x => x.Row).Select(x => new Expertise
            {
                Row = x.Key,
                Target = x.Sum(x => x.Target),
                Plan = x.Sum(x => x.Plan),
                Violation = x.Sum(x => x.Violation)
            }).ToArray();
            var specialists = reports.SelectMany(x => x.Specialists).GroupBy(x => x.Row).Select(x => new ZpzStatistics
            {
                Row = x.Key,
                Count = x.Sum(x => x.Count)
            }).ToArray();
            var complacence = reports.SelectMany(x => x.Complacence).GroupBy(x => x.Row).Select(x => new ZpzStatistics
            {
                Row = x.Key,
                Count = x.Sum(x => x.Count)
            }).ToArray();
            var informations = reports.SelectMany(x => x.Informations).GroupBy(x => x.Row).Select(x => new ZpzStatistics
            {
                Row = x.Key,
                Count = x.Sum(x => x.Count)
            }).ToArray();

            return new ZpzForWebSite
            {
                Filial = SummaryFilialCode,
                Treatments = treatments,
                Complacence = complacence,
                Complaints = complaints,
                Expertises = expertises,
                Informations = informations,
                Protections = protections,
                Specialists = specialists
            };
        }

        private ZpzForWebSite2023 CollectSummaryReport2023(ZpzForWebSite2023[] reports)
        {
            var treatments = reports.SelectMany(x => x.Treatments).GroupBy(x => x.Row).Select(x => new ZpzTreatment2023
            {
                Row = x.Key,
                Oral = x.Sum(x => x.Oral),
                Written = x.Sum(x => x.Written),
                Assignment = x.Sum(x => x.Assignment)
            }).ToArray();
            var complaints = reports.SelectMany(x => x.Complaints).GroupBy(x => x.Row).Select(x => new ZpzTreatment2023
            {
                Row = x.Key,
                Oral = x.Sum(x => x.Oral),
                Written = x.Sum(x => x.Written),
                Assignment = x.Sum(x => x.Assignment)

            }).ToArray();
            var protections = reports.SelectMany(x => x.Protections).GroupBy(x => x.Row).Select(x => new ZpzStatistics2023
            {
                Row = x.Key,
                Count = x.Sum(x => x.Count)
            }).ToArray();
            var expertises = reports.SelectMany(x => x.Expertises).GroupBy(x => x.Row).Select(x => new Expertise2023
            {
                Row = x.Key,
                Target = x.Sum(x => x.Target),
                Plan = x.Sum(x => x.Plan),
                Violation = x.Sum(x => x.Violation)
            }).ToArray();
            var specialists = reports.SelectMany(x => x.Specialists).GroupBy(x => x.Row).Select(x => new ZpzStatistics2023
            {
                Row = x.Key,
                Count = x.Sum(x => x.Count)
            }).ToArray();
            var informations = reports.SelectMany(x => x.Informations).GroupBy(x => x.Row).Select(x => new ZpzStatistics2023
            {
                Row = x.Key,
                Count = x.Sum(x => x.Count)
            }).ToArray();

            return new ZpzForWebSite2023
            {
                Filial = SummaryFilialCode,
                Treatments = treatments,
                Complaints = complaints,
                Expertises = expertises,
                Informations = informations,
                Protections = protections,
                Specialists = specialists
            };
        }

        private FFOMSMonthlyVol CollectSummaryFFOMSMonthlyVol(FFOMSMonthlyVol[] reports)
        {
            var SKP = reports.SelectMany(x => x.FFOMSMonthlyVol_SKP).GroupBy(x => x.RowNum).Select(x => new FFOMSMonthlyVol_SKP
            {
                RowNum = x.Key,
                CountSluch = x.Sum(x => x.CountSluch),
                CountAppliedSluch = x.Sum(x => x.CountAppliedSluch),
                CountSluchMEE = x.Sum(x => x.CountSluchMEE),
                CountSluchEKMP = x.Sum(x => x.CountSluchEKMP),
            }).ToArray();
            var SDP = reports.SelectMany(x => x.FFOMSMonthlyVol_SDP).GroupBy(x => x.RowNum).Select(x => new FFOMSMonthlyVol_SDP
            {
                RowNum = x.Key,
                CountSluch = x.Sum(x => x.CountSluch),
                CountAppliedSluch = x.Sum(x => x.CountAppliedSluch),
                CountSluchMEE = x.Sum(x => x.CountSluchMEE),
                CountSluchEKMP = x.Sum(x => x.CountSluchEKMP),
            }).ToArray();
            var APP = reports.SelectMany(x => x.FFOMSMonthlyVol_APP).GroupBy(x => x.RowNum).Select(x => new FFOMSMonthlyVol_APP
            {
                RowNum = x.Key,
                CountSluch = x.Sum(x => x.CountSluch),
                CountAppliedSluch = x.Sum(x => x.CountAppliedSluch),
                CountSluchMEE = x.Sum(x => x.CountSluchMEE),
                CountSluchEKMP = x.Sum(x => x.CountSluchEKMP),
            }).ToArray();
            var SMP = reports.SelectMany(x => x.FFOMSMonthlyVol_SMP).GroupBy(x => x.RowNum).Select(x => new FFOMSMonthlyVol_SMP
            {
                RowNum = x.Key,
                CountSluch = x.Sum(x => x.CountSluch),
                CountAppliedSluch = x.Sum(x => x.CountAppliedSluch),
                CountSluchMEE = x.Sum(x => x.CountSluchMEE),
                CountSluchEKMP = x.Sum(x => x.CountSluchEKMP),
            }).ToArray();

            return new FFOMSMonthlyVol
            {
                Filial = SummaryFilialCode,
                FFOMSMonthlyVol_SKP = SKP,
                FFOMSMonthlyVol_SDP = SDP,
                FFOMSMonthlyVol_APP = APP,
                FFOMSMonthlyVol_SMP = SMP
            };
        }

        private ViolationsOfAppeals CollectSummaryViolationsOfAppeals(ViolationsOfAppeals[] reports)
        {
            string yymm = reports.Select(x => x.Yymm).ElementAt(0);
            var table1 = reports.SelectMany(x => x.T1).GroupBy(x => x.Row).Select(x => new ForT1VOA
            {
                Row = x.Key,
                Oral = x.Sum(x => x.Oral),
                Written = x.Sum(x => x.Written),
                Assignment = x.Sum(x => x.Assignment)
            }).ToArray();
            var table2 = reports.SelectMany(x => x.T2).GroupBy(x => x.Row).Select(x => new ForT2VOA
            {
                Row = x.Key,
                Target = x.Sum(x => x.Target),
                Plan = x.Sum(x => x.Plan),
                Violation = x.Sum(x => x.Violation)
            }).ToArray();
            var table3 = reports.SelectMany(x => x.T3).GroupBy(x => x.Row).Select(x => new ForT3VOA
            {
                Row = x.Key,
                Target = x.Sum(x => x.Target),
                Plan = x.Sum(x => x.Plan),
                Violation = x.Sum(x => x.Violation)
            }).ToArray();

            return new ViolationsOfAppeals
            {
                Filial = SummaryFilialCode,
                Yymm = yymm,
                T1 = table1,
                T2 = table2,
                T3 = table3
            };
        }


        private FFOMSViolMEE CollectSummaryFFOMSViolMEE(FFOMSViolMEE[] reports)
        {
            var table1 = reports.SelectMany(x => x.DataViolMEE).GroupBy(x => x.RowNum).Select(x => new FFOMSViolMEEdata
            {
                RowNum = x.Key,
                Count = x.Sum(x => x.Count),
            }).ToArray();

            return new FFOMSViolMEE
            {
                Filial = SummaryFilialCode,
                DataViolMEE = table1,
            };
        }

        private FFOMSViolEKMP CollectSummaryFFOMSViolEKMP(FFOMSViolEKMP[] reports)
        {
            var table1 = reports.SelectMany(x => x.DataViolEKMP).GroupBy(x => x.RowNum).Select(x => new FFOMSViolEKMPdata
            {
                RowNum = x.Key,
                Count = x.Sum(x => x.Count),
            }).ToArray();

            return new FFOMSViolEKMP
            {
                Filial = SummaryFilialCode,
                DataViolEKMP = table1,
            };
        }

        private FFOMSVerifyPlan CollectSummaryFFOMSVerifyPlan(FFOMSVerifyPlan[] reports)
        {
            var table1 = reports.SelectMany(x => x.DataVerifyPlan).GroupBy(x => x.RowNum).Select(x => new FFOMSVerifyPlandata
            {
                RowNum = x.Key,
                Count = x.Sum(x => x.Count),
            }).ToArray();

            return new FFOMSVerifyPlan
            {
                Filial = SummaryFilialCode,
                DataVerifyPlan = table1,
            };
        }

        private FFOMSTargetedExp CollectSummaryFFOMSTargetedExp(FFOMSTargetedExp[] reports)
        {
            var table1 = reports.SelectMany(x => x.MEE).GroupBy(x => x.Row).Select(x => new MEE
            {
                Row = x.Key,
                Target = x.Sum(x => x.Target),

            }).ToArray();
            var table2 = reports.SelectMany(x => x.EKMP).GroupBy(x => x.Row).Select(x => new EKMP
            {
                Row = x.Key,
                Target = x.Sum(x => x.Target),
            }).ToArray();
            var table3 = reports.SelectMany(x => x.MD_EKMP).GroupBy(x => x.Row).Select(x => new MD_EKMP
            {
                Row = x.Key,
                Target = x.Sum(x => x.Target),
            }).ToArray();

            return new FFOMSTargetedExp
            {
                Filial = SummaryFilialCode,
                MEE = table1,
                EKMP = table2,
                MD_EKMP = table3
            };
        }


        private FFOMSPersonnel CollectSummaryFFOMSPersonnel(FFOMSPersonnel[] reports)
        {
            var table1 = reports.SelectMany(x => x.PersonnelT9).GroupBy(x => x.Row).Select(x => new PersonnelT9
            {
                Row = x.Key,
                FullTime = x.Sum(x => x.FullTime),
                Contract = x.Sum(x => x.Contract), 

            }).ToArray();

            return new FFOMSPersonnel
            {
                Filial = SummaryFilialCode,
                PersonnelT9 = table1,
            };
        }

        private ZpzForWebSite2025 CollectSummaryReport2025(ZpzForWebSite2025[] reports)
        {
            // Получаем все данные из WSData, преобразуем в массив
            var datas = reports.SelectMany(x => x.WSData).ToArray();
            string yymm = reports.Select(x => x.Yymm).ElementAt(0);
            // Считаем сумму по всем колонкам
            var summary = new WSData2025
            {
                Col1 = datas.Sum(x => x.Col1),
                Col2 = datas.Sum(x => x.Col2),
                Col3 = datas.Sum(x => x.Col3),
                Col4 = datas.Sum(x => x.Col4),
                Col5 = datas.Sum(x => x.Col5),
                Col6 = datas.Sum(x => x.Col6),
                Col8 = datas.Sum(x => x.Col8),
                Col9 = datas.Sum(x => x.Col9),
                Col10 = datas.Sum(x => x.Col10),
                Col11 = datas.Sum(x => x.Col11),
                Col12 = datas.Sum(x => x.Col12),
                Col13 = datas.Sum(x => x.Col13),
                Col14 = datas.Sum(x => x.Col14),
            };

            // Возвращаем итоговый отчет с суммами
            return new ZpzForWebSite2025
            {
                Filial = "Summary",  // Можно использовать любой идентификатор для итогового филиала
                Yymm = yymm,
                WSData = new WSData2025[] { summary }  // Возвращаем List, если нужно
            };
        }

        private void CreateReport(string filename, string filialName, FFOMSMonthlyVol report)
        {
            var excel = new ExcelFFOMSMonthlyVolCreator(filename, "", filialName);
            excel.CreateReport(report, null);
        }

        private void CreateReport(string filename, string filialName, ZpzForWebSite report)
        {
            var excel = new ExcelConsZpzWebSite(filename, "", filialName);
            excel.CreateReport(report, null);
        }

        private void CreateReport(string filename, string filialName, FFOMSViolMEE report)
        {
            var excel = new ExcelFFOMSViolMEECreator(filename, "", filialName);
            excel.CreateReport(report, null);
        }

        private void CreateReport(string filename, string filialName, FFOMSViolEKMP report)
        {
            var excel = new ExcelFFOMSViolEKMPCreator(filename, "", filialName);
            excel.CreateReport(report, null);
        }

        private void CreateReport(string filename, string filialName, FFOMSVerifyPlan report)
        {
            var excel = new ExcelFFOMSVerifyPlanCreator(filename, "", filialName);
            excel.CreateReport(report, null);
        }

        private void CreateReport(string filename, string filialName, FFOMSPersonnel report)
        {
            var excel = new ExcelFFOMSPersonnelCreator(filename, "", filialName);
            excel.CreateReport(report, null);
        }

        private void CreateReport(string filename, string filialName, ZpzForWebSite2023 report)
        {
            var excel = new ExcelConsZpzWebSite2023(filename, "", filialName);
            excel.CreateReport(report, null);
        }


        private void CreateReport(string filename, string filialName, ViolationsOfAppeals report)
        {
            var excel = new ExcelConsViolationsOfAppealsCreator(filename, "", filialName);
            excel.CreateReport(report, null);
        }

        private void CreateReport(string filename, string filialName, ZpzForWebSite2025 report)
        {
            var excel = new ExcelConsZpzWebSite2025(filename, "", filialName);
            excel.CreateReport(report, null);
        }

        private void CreateReport(string filename, string filialName, FFOMSTargetedExp report)
        {
            var excel = new ExcelFFOMSTargetedExpCreator(filename, "", filialName);
            excel.CreateReport(report, null);
        }



        private void CreateControlZpz(bool isMonthly)
        {
            string yymm = isMonthly ?
                GetYymm(cmbStart.Text, Convert.ToInt32(nudStart.Value)).ToString() :
                GetYymmQuarterly();

            var data = _client.CreateReportControlPgZpz(yymm, isMonthly);
            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (var d in data)
            {
                d.Filial = _regions.Single(x => x.Key == d.Filial).Value;
            }

            data = data.OrderBy(x => x.Filial).ToArray();

            string filename = saveFileDialog1.FileName;
            var excel = new ExcelControlZpzCreator(filename, "", _filialName, yymm);
            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(filename);
        }

        private void CreateControlZpz2023(bool isMonthly)
        {
            string yymm = isMonthly ?
                GetYymm(cmbStart.Text, Convert.ToInt32(nudStart.Value)).ToString() :
                GetYymmQuarterly();

            var data = _client.CreateReportControlZpz2023(yymm, isMonthly);
            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (var d in data)
            {
                d.Filial = _regions.Single(x => x.Key == d.Filial).Value;
            }

            data = data.OrderBy(x => x.Filial).ToArray();

            string filename = saveFileDialog1.FileName;
            var excel = new ExcelControlZpz2023Creator(filename, "", _filialName, yymm);
            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(filename);
        }

        private void CreateControlZpz2023Full()
        {
            string year = Convert.ToString(nudSingle.Value);
            var data = _client.CreateReportControlZpz2023Full(year);
            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (CurrentUser.IsMain)
            {
                foreach (var d in data)
                {
                    d.Filial = _regions.Single(x => x.Key == d.Filial).Value;
                }

                data = data.OrderBy(x => x.Filial).ToArray();

                string filename = saveFileDialog1.FileName;
                var excel = new ExcelControlZpz2023FullCreator(filename, "", _filialName);
                excel.CreateReport(data, null);

                GlobalUtils.OpenFileOrDirectory(filename);
            }
            else
            {
                foreach (var d in data)
                {
                    d.Filial = CurrentUser.Region;
                    data = data.OrderBy(x => x.Filial).ToArray();

                    string filename = saveFileDialog1.FileName;
                    var excel = new ExcelControlZpz2023FullCreator(filename, "", _filialName);
                    excel.CreateReport(data, null);

                    GlobalUtils.OpenFileOrDirectory(filename);
                }
            }
        }


        private void CreateControlZpz2025Full()
        {
            string year = Convert.ToString(nudSingle.Value);
            var data = _client.CreateReportControlZpz2025Full(year);
            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (CurrentUser.IsMain)
            {
                foreach (var d in data)
                {
                    d.Filial = _regions.Single(x => x.Key == d.Filial).Value;
                }

                data = data.OrderBy(x => x.Filial).ToArray();

                string filename = saveFileDialog1.FileName;
                var excel = new ExcelControlZpz2025FullCreator(filename, "", _filialName);
                excel.CreateReport(data, null);

                GlobalUtils.OpenFileOrDirectory(filename);
            }
            else
            {
                foreach (var d in data)
                {
                    d.Filial = CurrentUser.Region;
                    data = data.OrderBy(x => x.Filial).ToArray();

                    string filename = saveFileDialog1.FileName;
                    var excel = new ExcelControlZpz2025FullCreator(filename, "", _filialName);
                    excel.CreateReport(data, null);

                    GlobalUtils.OpenFileOrDirectory(filename);
                }
            }
        }


        private void CreateReportConsQuantityFilial()
        {
            string yymm = GetYymm(cmbStart.Text, Convert.ToInt32(nudStart.Value)).ToString();
            var dataMonths = _client.CreateReportConsQuantityFilial(yymm);
            if (dataMonths.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            foreach (var d in dataMonths)
            {
                if (d.Filial == "RU")
                {
                    continue;
                }
                else
                {
                    d.Filial = _regions.Single(j => j.Key == d.Filial).Value;
                }
            }
            dataMonths = dataMonths.OrderBy(x => x.Filial).Skip(1).ToArray();
            string statPeriod = yymm.Substring(0, 2) + "01";
            var dataYear = _client.CreateReportConsQuantityFilial(yymm);
            foreach (var d in dataYear)
            {
                if (d.Filial == "RU")
                {
                    continue;
                }
                else
                {
                    d.Filial = _regions.Single(k => k.Key == d.Filial).Value;
                }
            }

            dataYear = dataYear.OrderBy(x => x.Filial).Skip(1).ToArray();

            var excel = new ExcelConsolidateQuantityFilialsCreator(saveFileDialog1.FileName, "", _filialName);
            excel.CreateReport(dataMonths, dataYear);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }


        private void CreateFFOMSVolumesByTypes()
        {
            string yymm = GetYymmQuarterly();


            var data = _client.CreateFFOMSVolumesByTypes(yymm);

            if (data.VolFull.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var excel = new ExcelFFOMSVolumesByTypesCreator(saveFileDialog1.FileName, "", _filialName);
            excel.CreateReport(data,null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }

        private void CreateFFOMSLethalEKMP()
        {
            string yymm = GetYymmQuarterly();


            var data = _client.CollectFFOMSLethalEKMP(yymm);

            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var excel = new ExcelFFOMSLethalEKMPCreator(saveFileDialog1.FileName, "", _filialName);
            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }


        private void CreateZpzTable5()
        {
            string yymm = GetYymmQuarterly();


            var data = _client.CreateZpzTable5(yymm);

            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var excel = new ExcelZpzTable5Creator(saveFileDialog1.FileName, "", _filialName);
            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }


        private void CreateConsolidateQuantityAddRemove()
        {
            string year = nudSingle.Value.ToString();


            var data = _client.CreateConsolidateQuantityAddRemove(year);

            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var excel = new ExcelConsolidateQuantityAR(saveFileDialog1.FileName, "", _filialName);
            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }


        private void CreateConsolidateQuantityQ()
        {
            string yymm = GetYymmQuarterly();


            var data = _client.CreateConsolidateQuantityQ(yymm);

            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var excel = new ExcelConsolidateQuantityQ(saveFileDialog1.FileName, " " + yymm + " ", _filialName);
            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }


        private void CreateControlZpz2023Single()
        {
            string year = Convert.ToString(nudSingle.Value);
            string filial = CurrentUser.FilialCode;
            var data = _client.CreateReportControlZpz2023Single(year, filial);
            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (var d in data)
            {
                d.Filial = CurrentUser.Region;
                data = data.OrderBy(x => x.Filial).ToArray();

                string filename = saveFileDialog1.FileName;
                var excel = new ExcelControlZpz2023SingleCreator(filename, "", _filialName);
                excel.CreateReport(data, null);

                GlobalUtils.OpenFileOrDirectory(filename);
            }
        }


        private void CreateControlZpz2025Single()
        {
            string year = Convert.ToString(nudSingle.Value);
            string filial = CurrentUser.FilialCode;
            var data = _client.CreateReportControlZpz2025Single(year, filial);
            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (var d in data)
            {
                d.Filial = CurrentUser.Region;
                data = data.OrderBy(x => x.Filial).ToArray();

                string filename = saveFileDialog1.FileName;
                var excel = new ExcelControlZpz2025SingleCreator(filename, "", _filialName);
                excel.CreateReport(data, null);

                GlobalUtils.OpenFileOrDirectory(filename);
            }
        }


        private void CreateConsolidateQuantityFactPlan()
        {
            string year = nudSingle.Value.ToString();


            var data = _client.CreateConsolidateQuantityFactPlan(year);

            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var excel = new ExcelConsolidateQuantityFP(saveFileDialog1.FileName, "", _filialName);
            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }

        private void CreateConsolidateQuantityInfo()
        {
            string year = nudSingle.Value.ToString();


            var data = _client.CreateConsolidateQuantityInformation(year);

            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var excel = new ExcelConsolidateQuantityInfo(saveFileDialog1.FileName, "", _filialName);
            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }


        private void CreateOnko(bool isMonthly)
        {
            string yymm = isMonthly ?
                GetYymm(cmbStart.Text, Convert.ToInt32(nudStart.Value)).ToString() :
                GetYymmQuarterly();

            var data = _client.CreateConsolidateOnko(yymm, isMonthly);
            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (var d in data)
            {
                d.Filial = _regions.Single(x => x.Key == d.Filial).Value;
            }

            data = data.OrderBy(x => x.Filial).ToArray();

            string filename = saveFileDialog1.FileName;
            var excel = new ExcelOnkoCreator(filename, "", _filialName);
            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(filename);
        }

        private void CreateFull294()
        {
            string yymm = Convert.ToInt32(nudSingle.Value - 2000).ToString() + "12";

            var report = _client.CreateConsolidate294(yymm);
            //if (report != null && report.EfficiencyList != null && report.EfficiencyList.Length > 0)
            if (report != null)
            {
                var excel = new ExcelConsolidateFull294Creator(saveFileDialog1.FileName, "", _filialName);
                excel.CreateReport(report, null);
                GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
            }
            else
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void CreateFilial294()
        {
            var reportList = new List<Report294>();
            int endMonth = cmbStart.SelectedIndex;
            for (int i = 0; i <= endMonth; i++)
            {
                string filialCode = cmbRegion.SelectedValue.ToString();
                string month = GlobalConst.Months[i];
                string yymm = GetYymm(month, Convert.ToInt32(nudStart.Value)).ToString();

                var response = _client.GetReport(filialCode, yymm, ReportType.F294);
                var monthlyReport = response == null ? null : response as Report294;
                reportList.Add(monthlyReport);
            }

            if (reportList.Count > 0)
            {
                var excel = new ExcelConsolidateFilial294Creator(saveFileDialog1.FileName, "", _filialName);
                excel.CreateReport(reportList.ToArray(), null);
                GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
            }
            else
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void CreateReport262T3()
        {
            string yymm = GetYymm(cmbStart.Text, Convert.ToInt32(nudStart.Value)).ToString();
            var data = _client.CreateReport262T3(yymm);
            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (var d in data)
            {
                d.Filial = _regions.Single(x => x.Key == d.Filial).Value;
            }

            data = data.OrderBy(x => x.Filial).ToArray();

            var excel = new ExcelConsolidate262T3Creator(saveFileDialog1.FileName, "", _filialName);
            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }

        private void CreateReport262T2()
        {
            string yymm = GetYymm(cmbStart.Text, Convert.ToInt32(nudStart.Value)).ToString();
            var dataMonths = _client.CreateReport262T2(yymm, yymm);
            if (dataMonths.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (var d in dataMonths)
            {
                d.Filial = _regions.Single(x => x.Key == d.Filial).Value;
            }

            dataMonths = dataMonths.OrderBy(x => x.Filial).ToArray();

            string statPeriod = yymm.Substring(0, 2) + "01";
            var dataYear = _client.CreateReport262T2(statPeriod, yymm);
            foreach (var d in dataYear)
            {
                d.Filial = _regions.Single(x => x.Key == d.Filial).Value;
            }

            dataYear = dataYear.OrderBy(x => x.Filial).ToArray();

            var excel = new ExcelConsolidate262T2Creator(saveFileDialog1.FileName, "", _filialName);
            excel.CreateReport(dataMonths, dataYear);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }

        private void CreateReportCadreT2()
        {
            string yymm = GetYymm(cmbStart.Text, Convert.ToInt32(nudStart.Value)).ToString();
            var dataMonths = _client.CreateReportCadreTable2(yymm);
            if (dataMonths.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (var d in dataMonths)
            {
                if (d.Filial == "RU")
                {
                    continue;
                }
                else
                {
                    d.Filial = _regions.Single(j => j.Key == d.Filial).Value;
                }
            }

            dataMonths = dataMonths.OrderBy(x => x.Filial).Skip(1).ToArray();

            string statPeriod = yymm.Substring(0, 2) + "01";
            var dataYear = _client.CreateReportCadreTable2(yymm);
            foreach (var d in dataYear)
            {
                if (d.Filial == "RU")
                {
                    continue;
                }
                else
                {
                    d.Filial = _regions.Single(j => j.Key == d.Filial).Value;
                }
            }

            dataYear = dataYear.OrderBy(x => x.Filial).Skip(1).ToArray();

            var excel = new ExcelConsolidateCadreT2Creator(saveFileDialog1.FileName, "", _filialName);
            excel.CreateReport(dataMonths, dataYear);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }


        private void CreateReportCadreT1()
        {
            string yymm = GetYymm(cmbStart.Text, Convert.ToInt32(nudStart.Value)).ToString();
            var dataMonths = _client.CreateReportCadreTable1(yymm);
            if (dataMonths.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (var d in dataMonths)
            {
                if (d.Filial == "RU")
                {
                    continue;
                }
                else
                {
                    d.Filial = _regions.Single(j => j.Key == d.Filial).Value;
                }
            }

            dataMonths = dataMonths.OrderBy(x => x.Filial).Skip(1).ToArray();



            string statPeriod = yymm.Substring(0, 2) + "01";
            var dataYear = _client.CreateReportCadreTable1(yymm);
            foreach (var d in dataYear)
            {
                if (d.Filial == "RU")
                {
                    continue;
                }
                else
                {
                    d.Filial = _regions.Single(k => k.Key == d.Filial).Value;
                }
            }

            dataYear = dataYear.OrderBy(x => x.Filial).Skip(1).ToArray();

            var excel = new ExcelConsolidateCadreT1Creator(saveFileDialog1.FileName, "", _filialName);
            excel.CreateReport(dataMonths, dataYear);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }

        private void CreateReportVCRFilial()
        {
            string yymm = GetYymmQuarterly();
            var dataMonths = _client.CreateReportVCRFilial(yymm);
            if (dataMonths.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (var d in dataMonths)
            {
                if (d.Filial == "RU")
                {
                    continue;
                }
                else
                {
                    d.Filial = _regions.Single(j => j.Key == d.Filial).Value;
                }
            }

            dataMonths = dataMonths.OrderBy(x => x.Filial).ToArray();



            string statPeriod = yymm.Substring(0, 2) + "01";
            var dataYear = _client.CreateReportVCRFilial(yymm);
            foreach (var d in dataYear)
            {
                if (d.Filial == "RU")
                {
                    continue;
                }
                else
                {
                    d.Filial = _regions.Single(k => k.Key == d.Filial).Value;
                }
            }

            dataYear = dataYear.OrderBy(x => x.Filial).ToArray();

            var excel = new ExcelConsolidateVCRFilialCreator(saveFileDialog1.FileName, "", _filialName);
            excel.CreateReport(dataMonths, dataYear);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }


        private void CreateReportOpedUnplanned()
        {
            string yymm = GetYymm(cmbStart.Text, Convert.ToInt32(nudStart.Value)).ToString();
            var dataMonths = _client.CreateReportOpedUnplanned(yymm);
            if (dataMonths.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (var d in dataMonths)
            {
                if (d.Filial == "RU")
                {
                    continue;
                }
                else
                {
                    d.Filial = _regions.Single(j => j.Key == d.Filial).Value;
                }
            }

            dataMonths = dataMonths.OrderBy(x => x.Filial).ToArray();



            string statPeriod = yymm.Substring(0, 2) + "01";
            var dataYear = _client.CreateReportOpedUnplanned(yymm);
            foreach (var d in dataYear)
            {
                if (d.Filial == "RU")
                {
                    continue;
                }
                else
                {
                    d.Filial = _regions.Single(k => k.Key == d.Filial).Value;
                }
            }

            dataYear = dataYear.OrderBy(x => x.Filial).ToArray();

            var excel = new ExcelConsolidateOpedUnplannedCreator(saveFileDialog1.FileName, "", _filialName);
            excel.CreateReport(dataMonths, dataYear);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }

        private void CreateReport262T1()
        {
            int year = Convert.ToInt32(nudSingle.Value);
            var data = _client.CreateReport262T1(year);

            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (var d in data)
            {
                d.Filial = _regions.Single(x => x.Key == d.Filial).Value;
            }

            data = data.OrderBy(x => x.Filial).ToArray();

            var excel = new ExcelConsolidate262T1Creator(saveFileDialog1.FileName, "", _filialName);
            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }

        private void CreateCReportCpnp()
        {
            string yymm = GetYymmQuarterly();

            int q = cmbStart.SelectedIndex + 1;

            int year = Convert.ToInt32(nudStart.Value);


            var data = _client.CreateReportCpnp(yymm);

            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var excel = new ExcelConsolidateCnpnCreator(saveFileDialog1.FileName, "", _filialName, q, year);

            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);

        }


        private void CreateCReportCpnpMonth()
        {
            string yymm = GetYymm(cmbStart.Text, Convert.ToInt32(nudStart.Value)).ToString();


            int year = Convert.ToInt32(nudStart.Value);

            string mm = YymmUtils.GetMonth(yymm.Substring(2));

            var data = _client.CreateReportCpnpM(yymm);


            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var excel = new ExcelConsolidateCpnpMonthCreator(saveFileDialog1.FileName, "", _filialName, mm, year);

            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);

        }

        private void CreateCReportCardio()
        {
            string yymm = GetYymmQuarterly();

            var data = _client.CreateReportCardio(yymm);


            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var excel = new ExcelConsolidateCardio(saveFileDialog1.FileName, "", _filialName);

            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);

        }

        private void CreateCOped()
        {
            string yymmStart = GetYymm(cmbStart.Text, Convert.ToInt32(nudStart.Value)).ToString();
            string yymmEnd = GetYymm(cmbEnd.Text, Convert.ToInt32(nudEnd.Value)).ToString();
            ArrayOfString currentRegion = new ArrayOfString();
            if (cmbRegion.Text == "Все филиалы")
            {
                foreach (var region in _regions)
                {
                    currentRegion.Add(region.Key);

                }
            }
            else
            {
                currentRegion.Add(((KmsReportDictionary)cmbRegion.SelectedItem).Key);
            }


            var data = _client.CreateReportCOped(yymmStart, yymmEnd, currentRegion);

            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var excel = new ExcelConsolidateOpedCreator(saveFileDialog1.FileName, "", _filialName, yymmStart, yymmEnd);

            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);

        }

        private void CreateCReportDisp()
        {
            string yymm = GetYymmQuarterly();


            var data = _client.CreateReportDisp(yymm);


            if (data.Length == 0)
            {
                MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var excel = new ExcelConsolidateDisp(saveFileDialog1.FileName, "", _filialName);

            excel.CreateReport(data, null);

            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);

        }


        private void CreateCReportLetal()
        {
            string yymm = GetYymmQuarterly();

            if (Convert.ToInt32(yymm) < 2300)
            {
                var data = _client.CreateConsolidateLetal(yymm);
                if (data.Length == 0)
                {
                    MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                var excel = new ExcelConsolidateLetal(saveFileDialog1.FileName, "", _filialName);
                excel.CreateReport(data, null);
            }
            else
            {
                var data = _client.CreateConsolidateLetal2023(yymm);
                if (data.Length == 0)
                {
                    MessageBox.Show("По вашему запросу ничего не найдено", "Нет данных",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                var excel = new ExcelConsolidateLetal(saveFileDialog1.FileName, "", _filialName);
                excel.CreateReport(data, null);
            }
            GlobalUtils.OpenFileOrDirectory(saveFileDialog1.FileName);
        }


        private string GetYymmQuarterly()
        {
            decimal yy = nudStart.Value - 2000;
            string mmEnd = "0" + (3 * (Array.IndexOf(GlobalConst.Periods, cmbStart.Text) + 1)).ToString();

            return $"{yy}{mmEnd.Substring(mmEnd.Length - 2)}";
        }

        private string GetYymmQuarterly2()
        {
            decimal yy = nudStart.Value - 2000;
            string mmEnd = "0" + (3 * (Array.IndexOf(GlobalConst.PeriodsQ, cmbStart.Text) + 1)).ToString();

            return $"{yy}{mmEnd.Substring(mmEnd.Length - 2)}";
        }
    }
}