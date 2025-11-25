using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.ServiceModel;
using System.Windows.Forms;
using KmsReportClient.DgvHeaderGenerator;
using KmsReportClient.Excel.Collector;
using KmsReportClient.External;
using KmsReportClient.Forms.Dictionary;
using KmsReportClient.Global;
using KmsReportClient.Model;
using KmsReportClient.Model.Enums;
using KmsReportClient.Report;
using KmsReportClient.Report.Basic;
using KmsReportClient.Service;
using KmsReportClient.Support;
using NLog;
using static KmsReportClient.Global.GlobalConst;

namespace KmsReportClient.Forms
{
    public partial class MainForm : Form
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        StackedHeaderDecorator DgvRender;

        private readonly EndpointSoapClient _client;
        private readonly FileProcessor _ftpProcessor = new FileProcessor();
        private readonly ReportView _reportView;
        private readonly ExcelCollectorFactory _excelCollectorFactory = new ExcelCollectorFactory();

        private readonly Dictionary<string, IReportProcessor> _processorMap;
        private readonly Dictionary<TabPage, string> _tabControlMap;
        private readonly List<KmsReportDictionary> _regions;
        private readonly List<KmsReportDictionary> _reportsDictionary;



        public readonly string[] TreeTypes = { "Отчёты", "Запросы" };

        private string _currentReport;
        private string _currentReportName;


        private DynamicReportProcessor _dynamicReportProcessor;
        private int _currentDynamicReportFlow;
        private IReportProcessor _processor;
        private string _yymm;
        private bool _isQuery = false;

        public MainForm()
        {
            InitializeComponent();


            //PageElement.Visible = false;
            // SpravItem.Visible = false;
            // CmbTypeTree.Visible = false;

            var binding = new BasicHttpBinding
            {
                MaxReceivedMessageSize = 20 * 1024 * 1024,
                MaxBufferSize = 20 * 1024 * 1024,
                MaxBufferPoolSize = 20 * 1024 * 1024,
                SendTimeout = TimeSpan.FromSeconds(2000)
            };
            _client = new EndpointSoapClient();
            _client.Endpoint.Binding = binding;


            CheckUpdateApplication(true);

            _regions = CurrentUser.Regions.ToList();
            _regions.RemoveAt(39);
            _regions.Add(new KmsReportDictionary { Value = "Все филиалы", Key = "All" });
            _reportsDictionary = CurrentUser.ReportTypes.ToList();

            CmbFilterType.DisplayMember = "Value";
            CmbFilterType.ValueMember = "Key";
            CmbFilterType.DataSource = FilterList;
            CmbFilterType.SelectedIndex = 0;

            CmbFilials.DisplayMember = "Value";
            CmbFilials.ValueMember = "Key";
            CmbFilials.DataSource = _regions;

            _regions.RemoveAt(_regions.Count - 1);

            CmbStart.DisplayMember = "Value";
            CmbStart.ValueMember = "Key";
            CmbStart.DataSource = YymmUtils.GetMonths();

            CmbEnd.DisplayMember = "Value";
            CmbEnd.ValueMember = "Key";
            CmbEnd.DataSource = YymmUtils.GetMonths();



            _tabControlMap = CreateTabControlMap();
            _processorMap = CreateProcessorMap();

            _reportView = new ReportView(ReportTree, _regions, _reportsDictionary, _client);


            TreeYear.Maximum = DateTime.Today.Year;
            TreeYear.Value = DateTime.Today.Year;

            CmbTypeTree.DataSource = TreeTypes;

            DgvRender = new StackedHeaderDecorator(DgvQuery);
            _dynamicReportProcessor = new DynamicReportProcessor(_client, DgvQuery, CmbQuery);



            BtnUpload.Visible = false;
            Log.Info($"Старт работы формы. Пользователь {CurrentUser.UserName}");
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            MenuDownload.Visible = false;
            BtnCommentReport.Visible = false;
            BtnUploaded.Visible = false;
            BtnHandle.Visible = false;
            separatorExcel.Visible = true;
            PanelFilter.Enabled = true;

            TbControl.TabPages.Remove(PageIizl);
            TbControl.TabPages.Remove(Page262);
            TbControl.TabPages.Remove(Page294);
            TbControl.TabPages.Remove(PagePgQ);
            TbControl.TabPages.Remove(PagePg);
            TbControl.TabPages.Remove(PageQuery);
            TbControl.TabPages.Remove(PageOped);
            TbControl.TabPages.Remove(PageOpedU);
            TbControl.TabPages.Remove(PageOpedQ);
            TbControl.TabPages.Remove(PageOtclkInfrorm);
            TbControl.TabPages.Remove(tabVac);
            TbControl.TabPages.Remove(PageFssMonitoring);
            TbControl.TabPages.Remove(PageMonitoringVCR);
            TbControl.TabPages.Remove(PageIizl);
            TbControl.TabPages.Remove(PageProposal);
            TbControl.TabPages.Remove(tpOpedFinance);
            TbControl.TabPages.Remove(tpOpedFinance3);
            TbControl.TabPages.Remove(tpIizl2022);
            TbControl.TabPages.Remove(PageCadre);
            TbControl.TabPages.Remove(PageReqVCR);
            TbControl.TabPages.Remove(PageZpz);
            TbControl.TabPages.Remove(PageZpzQ);
            TbControl.TabPages.Remove(PageZpz10);
            TbControl.TabPages.Remove(PageZpzLethal);
            TbControl.TabPages.Remove(PageEffectiveness);
            TbControl.TabPages.Remove(PageQuantity);
            TbControl.TabPages.Remove(PageTarAllow);
            TbControl.TabPages.Remove(PageDoff);
            TbControl.TabPages.Remove(PageZpz2025);
            TbControl.TabPages.Remove(PageZpzQ2025);
            TbControl.TabPages.Remove(PageZpz10_2025);
            TbControl.TabPages.Remove(PageZpzLethal2025);
            TbControl.TabPages.Remove(PageViolMEE);
            TbControl.TabPages.Remove(PageViolEKMP);
            TbControl.TabPages.Remove(PageVerifyPlan);
            TbControl.TabPages.Remove(PageMonthlyVol);

            if (CurrentUser.IsMain)
            {
                CreateNewFilter();
                SetVisibilityHeadOffice();
            }
            else
            {
                SetVisibilityFilials();
            }

            _currentReport = "";
        }

        private Dictionary<TabPage, string> CreateTabControlMap() =>
            new Dictionary<TabPage, string> {
                        {Page262, ReportGlobalConst.Report262},
                        {Page294, ReportGlobalConst.Report294},
                        {PageIizl, ReportGlobalConst.ReportIizl},
                        {PagePg, ReportGlobalConst.ReportPg},
                        {PagePgQ, ReportGlobalConst.ReportPgQ},
                        {PageQuery, ReportGlobalConst.ReportPgQ},
                        {PageZpz, ReportGlobalConst.ReportZpz},
                        {PageZpzQ, ReportGlobalConst.ReportZpzQ},
                        {PageQueryZpz, ReportGlobalConst.ReportZpzQ},
                        {PageZpz2025, ReportGlobalConst.ReportZpz2025},
                        {PageZpzQ2025, ReportGlobalConst.ReportZpzQ2025},
                        {PageQueryZpz2025, ReportGlobalConst.ReportZpzQ2025},
                        {PageOped, ReportGlobalConst.ReportOped},
                        {PageOpedU, ReportGlobalConst.ReportOpedU},
                        {PageOtclkInfrorm, ReportGlobalConst.ReportOtklik},
                        {PageOpedQ, ReportGlobalConst.ReportOpedQ},
                        {tabVac, ReportGlobalConst.ReportVac},
                        {PageFssMonitoring, ReportGlobalConst.FSSMonitoring},
                        {PageMonitoringVCR, ReportGlobalConst.MonitoringVCR},
                        {PageProposal, ReportGlobalConst.ReportProposal},
                        {tpOpedFinance, ReportGlobalConst.ReportOpedFinance},
                        {tpOpedFinance3, ReportGlobalConst.ReportOpedFinance},
                        {tpIizl2022, ReportGlobalConst.ReportIizl2022},
                        {PageCadre, ReportGlobalConst.ReportCadre},
                        {PageReqVCR, ReportGlobalConst.ReportReqVCR},
                        {PageZpz10, ReportGlobalConst.ReportZpz10},
                        {PageZpzLethal, ReportGlobalConst.ReportZpzLethal},
                        {PageZpz10_2025, ReportGlobalConst.ReportZpz10_2025},
                        {PageZpzLethal2025, ReportGlobalConst.ReportZpz2025Lethal},
                        {PageEffectiveness, ReportGlobalConst.ReportEffectiveness},
                        {PageQuantity, ReportGlobalConst.ReportQuantity},
                        {PageTarAllow, ReportGlobalConst.ReportTargetedAllowances},
                        {PagePVPLoad, ReportGlobalConst.ReportPVPLoad},
                        {PageDoff, ReportGlobalConst.ReportDoff},
                        {PageViolMEE, ReportGlobalConst.ReportViolMEE},
                        {PageViolEKMP, ReportGlobalConst.ReportViolEKMP},
                        {PageVerifyPlan, ReportGlobalConst.ReportVerifyPlan},
                        {PageMonthlyVol, ReportGlobalConst.ReportMonthlyVol},
            };

        private Dictionary<string, IReportProcessor> CreateProcessorMap() =>
            new Dictionary<string, IReportProcessor>
            {
                {
                    ReportGlobalConst.Report262,
                    new Report262Processor(_client, _reportsDictionary, DgwReport262, Cmb262, Txtb262, Page262)
                },
                {
                    ReportGlobalConst.Report294,
                    new Report294Processor(_client, _reportsDictionary, DgwReport294, Cmb294, Txtb294, Page294)
                },
                {
                    ReportGlobalConst.ReportIizl,
                    new ReportIizlProcessor(_client, _reportsDictionary, DgwReportIizl, CmbIizl, TxtbIizl, PageIizl)
                },
                {
                    ReportGlobalConst.ReportPg,
                    new ReportPgProcessor(_client, _reportsDictionary, DgwReportPg, CmbPg, TxtbPg, PagePg)
                },
                {
                    ReportGlobalConst.ReportPgQ,
                    new ReportPgQProcessor(_client, _reportsDictionary, DgwReportPgQ, CmbPgQ, TxtbPgQ, PagePgQ)
                },
                {
                    ReportGlobalConst.ReportZpz,
                    new ReportZpzProcessor(_client, _reportsDictionary, DgwReportZpz, CmbZpz, TxtbZpz, PageZpz)
                },
                {
                    ReportGlobalConst.ReportZpzLethal,
                    new ReportZpzLethalProcessor(_client, _reportsDictionary, DgwReportZpzLethal, CmbZpzLethal, TxtbZpzLethal, PageZpzLethal)
                },
                {
                    ReportGlobalConst.ReportZpz10,
                    new ReportZpz10Processor(_client, _reportsDictionary, DgwReportZpz10, CmbZpz10, TxtbZpz10, PageZpz10)
                },
                {
                    ReportGlobalConst.ReportZpzQ,
                    new ReportZpzQProcessor(_client, _reportsDictionary, DgwReportZpzQ, CmbZpzQ, TxtbZpzQ, PageZpzQ)
                },
                                {
                    ReportGlobalConst.ReportZpz2025,
                    new ReportZpz2025Processor(_client, _reportsDictionary, DgvReportZpz2025, CmbZpz2025, TxtbZpz2025, PageZpz2025)
                },
                {
                    ReportGlobalConst.ReportZpz2025Lethal,
                    new ReportZpz2025LethalProcessor(_client, _reportsDictionary, DgvReportZpzLethal2025, CmbZpzLethal2025, TxtbZpzLethal2025, PageZpzLethal2025)
                },
                {
                    ReportGlobalConst.ReportZpz10_2025,
                    new ReportZpz10_2025Processor(_client, _reportsDictionary, DgvReportZpz10_2025, CmbZpz10_2025, TxtbZpz10_2025, PageZpz10_2025)
                },
                {
                    ReportGlobalConst.ReportZpzQ2025,
                    new ReportZpzQ2025Processor(_client, _reportsDictionary, DgvReportZpzQ2025, CmbZpzQ2025, TxtbZpzQ2025, PageZpzQ2025)
                },
                 {
                    ReportGlobalConst.ReportOped,
                    new ReportOpedProcessor(_client, _reportsDictionary, DgvReportOped, CbxOped, TxtbOped, PageOped)
                },
                 {
                    ReportGlobalConst.ReportOpedU,
                    new ReportOpedUProcessor(_client, _reportsDictionary, DgvReportOpedU, CbxOpedU, TxtbOpedU, PageOpedU)
                },
                  {
                    ReportGlobalConst.ReportOpedQ,
                    new ReportOpedQProcessor(_client, _reportsDictionary, dgvOpedQ, cmbOpedQ, tbOpedQ, PageOpedQ)
                },
                  {
                    ReportGlobalConst.ReportEffectiveness,
                    new ReportEffectivenessProcessor(_client, _reportsDictionary, dgvEffectiveness, cmbEffectiveness, tbEffectiveness, PageEffectiveness)
                },
                  {
                    ReportGlobalConst.ReportOtklik,
                    new ReportInfrormationResponseProcessor(_client, _reportsDictionary, DgvOtclkInfrorm, CbxOtclkInfrorm, TxtOtclkInfrorm, PageOtclkInfrorm)
                },
                  {
                    ReportGlobalConst.ReportVac,
                    new ReportVaccinationProccesor(_client, _reportsDictionary, gVac, cbVac, tbVac, tabVac)
                },
                  {
                    ReportGlobalConst.FSSMonitoring,
                    new FSSMonitoringProcessor(_client, _reportsDictionary, dgvFssM, cbFssM, tbFssM, PageFssMonitoring)
                },
                  {
                    ReportGlobalConst.MonitoringVCR,
                    new MonitoringVCRProcessor(_client, _reportsDictionary, dgvMonitoringVCR, cbMonitoringVCR, tbMonitoringVCR, PageMonitoringVCR)
                },

                  {
                    ReportGlobalConst.ReportProposal,
                    new ReportProposalProcessor(_client, _reportsDictionary, dgvProposal, cbProposal, tbProposal, PageProposal)
                },

                   {
                    ReportGlobalConst.ReportOpedFinance,
                    new ReportOpedFinanceProcessor(_client, _reportsDictionary, dgvOpedFinance, cbOpedFinance, tbOpedFinance, tpOpedFinance)
                },

                                   {
                    ReportGlobalConst.ReportOpedFinance3,
                    new ReportOpedFinance3Processor(_client, _reportsDictionary, dgvOpedFinance3, cbOpedFinance3, tbOpedFinance3, tpOpedFinance3)
                },

                  {
                    ReportGlobalConst.ReportIizl2022,
                    new ReportIizlProcessor2022(_client, _reportsDictionary, dgvIizl2022, cbIizl2022, tbIizl2022, tpIizl2022)
                },
                {
                    ReportGlobalConst.ReportCadre,
                    new ReportCadreProcessor(_client, _reportsDictionary, DgvCadre, CmbCadre, TxtbCadre, PageCadre)
                },
                {
                    ReportGlobalConst.ReportReqVCR,
                    new ReportReqVCRProcessor(_client, _reportsDictionary, DgvReqVCR, CmbReqVCR, TxtbReqVCR, PageReqVCR)
                },
                {
                    ReportGlobalConst.ReportQuantity,
                    new ReportQuantityProcessor(_client, _reportsDictionary, DgvQuantity, CmbQuantity, TxtbQuantity, PageQuantity)
                },
                {
                    ReportGlobalConst.ReportTargetedAllowances,
                    new ReportTargetedAllowancesProcessor(_client, _reportsDictionary, DgvTarAllow, CmbTarAllow, TbTarAllow, PageTarAllow)
                },
                {
                    ReportGlobalConst.ReportPVPLoad,
                    new ReportPVPLoadProcessor(_client, _reportsDictionary, DgvPVPLoad, CmbPVPLoad, TbPVPLoad, PagePVPLoad)
                },
                {
                    ReportGlobalConst.ReportDoff,
                    new ReportDoffProcessor(_client, _reportsDictionary, DgvDoff, CmbDoff, TbDoff, PageDoff)
                },
                {
                    ReportGlobalConst.ReportViolMEE,
                    new ReportViolMEEProcessor(_client, _reportsDictionary, DgvViolMEE, CmbViolMEE, TbViolMEE, PageViolMEE)
                },
                {
                    ReportGlobalConst.ReportVerifyPlan,
                    new ReportVerifyPlanProcessor(_client, _reportsDictionary, DgvVerifyPlan, CmbVerifyPlan, TbVerifyPlan, PageVerifyPlan)
                },
                {
                    ReportGlobalConst.ReportViolEKMP,
                    new ReportViolEKMPProcessor(_client, _reportsDictionary, DgvViolEKMP, CmbViolEKMP, TbViolEKMP, PageViolEKMP)
                },
                {
                    ReportGlobalConst.ReportMonthlyVol,
                    new ReportMonthlyVolProcessor(_client, _reportsDictionary, DgvMonthlyVol, CmbMonthlyVol, TbMonthlyVol, PageMonthlyVol)
                }
            };

        private void CreateNewFilter()
        {
            CmbEnd.SelectedIndex = DateTime.Today.Month - 1;
            CmbStart.SelectedIndex = DateTime.Today.Month - 1;
            CmbFilials.SelectedIndex = 0;
            CmbFilterType.SelectedIndex = 0;

            TxtbFilials.Clear();
            NumStart.Value = DateTime.Today.Year;
            NumEnd.Value = DateTime.Today.Year;
        }

        private void SetVisibilityHeadOffice()
        {
            BtnOpen.Visible = false;
            BtnSave.Visible = false;
            BtnClear.Visible = false;
            BtnUpload.Visible = false;
            BtnUploaded.Visible = false;
            BtnAutoFill.Visible = false;
            BtnSummary.Visible = false;
            BtnSubmit.Text = "Утвердить отчет";
            consolidateMenu.Visible = true;
            serviceMenu.Visible = true;
            BtnSaveToDb.Visible = false;
            separatorExcel.Visible = false;
            toolStripSeparator2.Visible = false;
        }

        private void SetVisibilityFilials()
        {
            TbControl.Dock = DockStyle.Fill;
            PanelFilter.Visible = false;
            consolidateMenu.Visible = false;
            serviceMenu.Visible = false;
            BtnRefuse.Visible = false;
            SpravItem.Visible = false;
            Con.Visible = false;
            BtnSummary.Visible = false;
        }

        private void CreateTreeView()
        {

            CmbTypeTree.SelectedIndex = 0;
            bool isNeedRefuseNotification = _reportView.CreateTreeView((int)TreeYear.Value);
            if (isNeedRefuseNotification && !CurrentUser.IsMain)
            {
                MessageBox.Show(
                                    "Имеются отчеты, возвращенные на доработку",
                                    "Внимание",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning
                               );
            }
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Maximized;
            TsslVersion.Width = 100;
            TsslVersion.Text = $"Версия: {Application.ProductVersion}";
            Text = $"Регион: {CurrentUser.Region}. Пользователь: {CurrentUser.UserName}";
        }

        private void CreateReportTab()
        {
            NotifyAboutEnabledFilter();

            bool isNeedCreateReport = false;
            string filialCode = "";

            if (CurrentUser.IsMain && ReportTree.SelectedNode.Level == 3)
            {
                isNeedCreateReport = true;
                _yymm = ReportTree.SelectedNode.Parent.Text;
                _currentReportName = ReportTree.SelectedNode.Parent.Parent.Text;

                filialCode = _regions.Single(x => x.Value == ReportTree.SelectedNode.Text).Key;


            }
            else if (!CurrentUser.IsMain && ReportTree.SelectedNode.Level == 2)
            {
                isNeedCreateReport = true;
                _yymm = ReportTree.SelectedNode.Text;
                _currentReportName = ReportTree.SelectedNode.Parent.Text;

                filialCode = CurrentUser.FilialCode;

            }
            //else if (!CurrentUser.IsMain && (ReportTree.SelectedNode.Text == "Нарушения МЭЭ" || ReportTree.SelectedNode.Text == "Нарушения ЭКМП" || ReportTree.SelectedNode.Text == "Планы проверок" || ReportTree.SelectedNode.Text == "Объемы ежемесячные"))
            //{
            //    isNeedCreateReport = true;
            //    _yymm = "Март 2025";
            //    _currentReportName = ReportTree.SelectedNode.Text;

            //    filialCode = CurrentUser.FilialCode;
            //}
            //else if (CurrentUser.IsMain)
            //{
            //    try
            //    {
            //        if (ReportTree.SelectedNode.Parent.Text == "Нарушения МЭЭ" ||
            //    ReportTree.SelectedNode.Parent.Text == "Нарушения ЭКМП" ||
            //    ReportTree.SelectedNode.Parent.Text == "Планы проверок" ||
            //    ReportTree.SelectedNode.Parent.Text == "Объемы ежемесячные")
            //        {
            //            isNeedCreateReport = true;
            //            _yymm = "Март 2025";
            //            _currentReportName = ReportTree.SelectedNode.Parent.Text;
            //            try
            //            {
            //                filialCode = _regions.Single(x => x.Value == ReportTree.SelectedNode.Text).Key;
            //            }
            //            catch (Exception ex) { }

            //        }
            //    }
            //    catch (Exception ex) { }
            //}
            Console.WriteLine($"yymm={_yymm} currentReportName={_currentReportName} Филиал={filialCode}");
            
            if (isNeedCreateReport)
            {
                _currentReport = _reportsDictionary.Single(x => x.Value == _currentReportName).Key;
                _processor = _processorMap[_currentReport];
                _processor.FilialCode = filialCode;
                OpenReport();
            }

        }


        private void CreateDynamicReportTab()
        {
            NotifyAboutEnabledFilter();
            string filialCode = "";

            bool isNeedCreateReport = false;
            if (CurrentUser.IsMain && (ReportTree.SelectedNode.Level == 2 || ReportTree.SelectedNode.Level == 1))
            {
                isNeedCreateReport = true;
                if (ReportTree.SelectedNode.Level != 1)
                {
                    _yymm = ReportTree.SelectedNode.Parent.Parent.Text;
                }
                else
                {
                    _yymm = ReportTree.SelectedNode.Parent.Text;

                }

                _currentReportName = ReportTree.SelectedNode.Parent.Text;
                _currentDynamicReportFlow = (ReportTree.SelectedNode.Tag as ReportNodeTag).idFlow;
                //filialCode = _regions.Single(x => x.Value == ReportTree.SelectedNode.Text).Key;
            }
            else if (!CurrentUser.IsMain && ReportTree.SelectedNode.Level == 1)
            {
                isNeedCreateReport = true;
                _yymm = ReportTree.SelectedNode.Parent.Text;
                _currentReportName = ReportTree.SelectedNode.Text;
                _currentDynamicReportFlow = (ReportTree.SelectedNode.Tag as ReportNodeTag).idFlow;
                filialCode = CurrentUser.FilialCode;
            }



            if (isNeedCreateReport)
            {
                OpenDynamicReport();
            }
        }


        private void OpenDynamicReport()
        {
            if (!TbControl.TabPages.Contains(PageQuery))
            {
                TbControl.TabPages.Add(PageQuery);
            }
            _dynamicReportProcessor.data.Clear();
            var reportTag = ReportTree.SelectedNode.Tag as ReportNodeTag;
            var reportResponse = _dynamicReportProcessor.GetXmlReport(reportTag.IdReport);
            DgvQuery.Rows.Clear();
            DgvQuery.Columns.Clear();
            _dynamicReportProcessor.SetReport(reportResponse);
            _dynamicReportProcessor.SetReportDynamic(reportTag.IdReport);
            _dynamicReportProcessor.SetComboBox(CmbQuery);
            _dynamicReportProcessor.SetReadOnlyDgv(DgvQuery, reportTag.idFlow);
            TxtbInfo.Text = _dynamicReportProcessor.GetReportInfo(reportTag.idFlow);
            if (reportTag.idFlow != 0)
            {
                var data = _dynamicReportProcessor.GetRegionData(reportTag.idFlow);
                _dynamicReportProcessor.data = data;
                // пробуем всунуть создание dgv
                _dynamicReportProcessor.SetDgv(DgvQuery, CmbQuery.Text);
                _dynamicReportProcessor.SetReportDynamicFlow(reportTag.idFlow);
                _dynamicReportProcessor.FillThemeData(DgvQuery);
            }
        }

        private void NotifyAboutEnabledFilter()
        {
            if (!ChkbFilter.Checked)
            {
                return;
            }

            var dialogResult = MessageBox.Show(
                "Окно фильтрации активно. Закрыть окно фильтра и открыть отчет филиала?",
                "Вопрос",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (dialogResult == DialogResult.No)
            {
                return;
            }

            ChkbFilter.Checked = false;
        }

        private void OpenReport()
        {
            var yymmExp = YymmUtils.ConvertPeriodToYymm(_yymm);
            var inReport = _processor.CollectReportFromWs(yymmExp);
            _processor.OldTheme = _processor.HasReport ? _processor.GetCurrentTheme() : _processor.ThemesList[0].Key;

            if (!_processor.HasReport)
            {
                TbControl.TabPages.Add(_processor.Page);
            }

            if (inReport != null)
            {
                _processor.Report = inReport;
            }
            else
            {
                _processor.InitReport();
            }

            _processor.HasReport = true;
            _processor.Report.Yymm = YymmUtils.ConvertPeriodToYymm(_yymm);
            _processor.ColorReport = ReportTree.SelectedNode.BackColor;
            _processor.CreateReportForm(_processor.OldTheme);
            _processor.FillDataGridView(_processor.OldTheme);
            _processor.SetReadonlyForDgv(SuccessStatuses.Contains(_processor.Report.Status));

            SetReportInterface();

            if (CurrentUser.IsMain && inReport == null)
            {
                var dialogResult = MessageBox.Show("Филиал еще не вносил данные по выбранному периоду",
                    "Информация",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.ServiceNotification
                    );
            }
        }


        private void DgvReportOped_CellValueChanged1(object sender, DataGridViewCellEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void DgvReportOpedU_CellValueChanged1(object sender, DataGridViewCellEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void SetReportInterface()
        {
            bool isEnabled = !SuccessStatuses.Contains(_processor.Report.Status);

            BtnOpen.Enabled = isEnabled;
            BtnSave.Enabled = isEnabled;
            BtnClear.Enabled = isEnabled;
            BtnUpload.Enabled = isEnabled;
            BtnAutoFill.Enabled = isEnabled;
            BtnSaveToDb.Enabled = isEnabled;

            BtnSubmit.Enabled = _processor.Report.Status != ReportStatus.Done;
            Console.WriteLine(!CurrentUser.IsMain);

            BtnUploaded.Visible = _processor.IsVisibleBtnDownloadExcel();
            BtnHandle.Visible = _processor.IsVisibleBtnHandle();
            BtnSummary.Visible = _processor.IsVisibleBtnSummary();
            if (_processor.Report.IdType == "TarAllow")
            {
                BtnSubmit.Visible = false;
            }
            if (_processor.Report.IdType != "TarAllow")
            {
                BtnSubmit.Visible = true;
            }
            if (_isQuery)
            {
                BtnUploaded.Visible = false;
                BtnHandle.Visible = false;
                BtnSummary.Visible = false;

            }

            if (_processor.Report.IdType == "PG" || _processor.Report.IdType == "Zpz" || _processor.Report.IdType == "Zpz2025")
            {
                DgwReportPg.ReadOnly = _processor.Report.DataSource != DataSource.Handle;
                DgwReportZpz.ReadOnly = _processor.Report.DataSource != DataSource.Handle;
                DgvReportZpz2025.ReadOnly = _processor.Report.DataSource != DataSource.Handle;
            }
            TxtbInfo.Text = _processor.GetReportInfo();
            BtnCommentReport.Visible = true;
            BtnCommentReport.DisplayStyle = CheckComment()
                ? ToolStripItemDisplayStyle.ImageAndText
                : ToolStripItemDisplayStyle.Text;

            TbControl.SelectedTab = _processor.Page;
            Console.WriteLine(_processor.FilialCode);
            TbControl.SelectedTab.Text = !CurrentUser.IsMain ? $"{_currentReportName} | {_yymm}"
                : ChkbFilter.Checked ? $"{_currentReportName} | Фильтр"
                : $"{_currentReportName} | {_yymm} | {_regions.SingleOrDefault(x => x.Key == _processor.FilialCode).Value}";
        }


        private void DeserializeReport()
        {
            if (string.IsNullOrEmpty(_currentReport))
            {
                return;
            }

            var oldTheme = _processor.GetCurrentTheme();
            _processor.DeserializeReport(YymmUtils.ConvertPeriodToYymm(_yymm));
            _processor.OldTheme = oldTheme;
            _processor.CreateReportForm(oldTheme);
            _processor.FillDataGridView(oldTheme);
        }

        private void OpenScan()
        {
            if (string.IsNullOrEmpty(_currentReport))
            {
                return;
            }

            try
            {
                if (string.IsNullOrEmpty(_processor.Report.Scan))
                {
                    throw new Exception("Скан отчета не загружен на сервер!");
                }

                string downloadFilename = _ftpProcessor
                    .DownloadFileFromWs(_processor.Report.Scan, "", _processor.FilialCode, _client);
                Process.Start(downloadFilename);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка открытия скана");
                MessageBox.Show("Ошибка открытия скана: " + ex.Message, "Ошибка", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void SerializeReport()
        {
            _processor.MapReportFromDgv(_processor.GetCurrentTheme());
            _processor.Serialize(YymmUtils.ConvertPeriodToYymm(_yymm));
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_currentReport))
            {
                return;
            }

            var dialogResult = MessageBox.Show("Вы уверены, что хотите очистить текущую форму?",
                "Очистить форму?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
                _processor.CreateReportForm(_processor.OldTheme);
            }
        }

        private void SaveDynamycReportToDb()
        {
            var flow = _dynamicReportProcessor.GetReportDynamicFlow(_currentDynamicReportFlow);
            if (flow != null)
            {
                if (flow.Status == ReportStatus.Submit)
                {
                    return;
                }
            }

            _dynamicReportProcessor.SetData(DgvQuery, _dynamicReportProcessor._pageIndex);
            int flowid = _dynamicReportProcessor.SaveReportFiliialData();
            if (flowid != 0)
            {
                ReportTree.SelectedNode.BackColor = ColorBd;
                (ReportTree.SelectedNode.Tag as ReportNodeTag).idFlow = flowid;
            }
        }

        private void SaveReportToDb()
        {
            if (string.IsNullOrEmpty(_currentReport))
            {
                return;
            }

            bool isNotNeedSave = CurrentUser.IsMain || SuccessStatuses.Contains(_processor.Report.Status);
            if (isNotNeedSave)
            {
                return;
            }

            try
            {
                SerializeReport();
                var message = _processor.ValidReport();

                if (message.Length > 0)
                {
                    TxtbInfo.Text = message;
                    MessageBox.Show(@"В отчете имеются ошибки. Перед сохранием их необходимо исправить",
                        @"Контроль формы", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                _processor.SaveToDb();
                if (_processor.Report.Status != ReportStatus.Scan)
                {
                    ReportTree.SelectedNode.BackColor = ColorBd;
                    _processor.Report.Status = ReportStatus.Saved;
                }

                MessageBox.Show("Отчет успешно сохранен на сервере", "Сохранение отчета!", MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка сохранения отчета в Базе данных");
                MessageBox.Show("Ошибка сохранения отчета в Базе данных: " + ex, "Ошибка", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                //Console.WriteLine(ex.StackTrace);
            }
        }

        private void SaveReportToDB_DataSourceToExcel()
        {
            if (string.IsNullOrEmpty(_currentReport))
            {
                return;
            }

            if (_processor.Report.DataSource != DataSource.Excel)
            {
                _processor.Report.DataSource = DataSource.Excel;
            }
            _processor.SaveReportDataSourceExcel();
            DgwReportPg.ReadOnly = true;
            DgwReportZpz.ReadOnly = true;
            DgwReportZpz10.ReadOnly = false;
            DgvReportZpz2025.ReadOnly = true;
            DgvReportZpz10_2025.ReadOnly = false;
            _processor.SaveToDb();
        }

        private void SaveReportToDB_DataSourceToHandle()
        {
            string theme = _processor.GetCurrentTheme();
            if (string.IsNullOrEmpty(_currentReport))
            {
                return;
            }

            if (_processor.Report.DataSource != DataSource.Handle)
            {
                _processor.Report.DataSource = DataSource.Handle;
            }

            _processor.SaveReportDataSourceHandle();
            DgwReportPg.ReadOnly = false;
            DgwReportZpz.ReadOnly = false;
            DgwReportZpz10.ReadOnly = false;
            DgvReportZpz2025.ReadOnly = false;
            DgvReportZpz10_2025.ReadOnly = false;
            _processor.FillDataGridView(theme);
            _processor.SaveToDb();

        }

        private void UploadToExcel()
        {
            bool emptyReport = string.IsNullOrEmpty(_currentReport);
            if (emptyReport)
            {
                return;
            }
            else
            {
                try
                {
                    saveFileDialog1.FileName = ChkbFilter.Checked
                        ? $"Сводный_отчет_{_processor.SmallName}_{_processor.Report.Yymm}.xlsx"
                        : GetFileName(".xlsx");
                    string reportFilialName = CurrentUser.IsMain && ChkbFilter.Checked ? "ООО \"Капитал МС\"" :
                        _processor.FilialName;

                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        _processor.MapReportFromDgv(_processor.GetCurrentTheme());
                        _processor.ToExcel(saveFileDialog1.FileName, reportFilialName);
                        OpenFileOrDirectory(saveFileDialog1.FileName);
                    }
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "Ошибка выгрузки документа в Excel");
                    MessageBox.Show("Ошибка выгрузки документа в Excel: " + ex, "Ошибка", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }

        }

        private void UploadToExcelDynamicReport()
        {
            var tag = ReportTree.SelectedNode.Tag as ReportNodeTag;

            if (tag.IdReport == 0)
            {
                return;
            }

            try
            {
                var report = _dynamicReportProcessor.GetReportDynamic(tag.IdReport);
                saveFileDialog1.FileName = String.Format($"{report.Name.Trim()}.xlsx");




                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    if (tag.idFlow != 0)
                    {
                        var data = _dynamicReportProcessor.GetRegionData(tag.idFlow);
                        _dynamicReportProcessor.ToExcel(saveFileDialog1.FileName, data);
                    }
                    else
                    {
                        _dynamicReportProcessor.ToExcel(saveFileDialog1.FileName);
                    }

                    OpenFileOrDirectory(saveFileDialog1.FileName);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка выгрузки документа в Excel");
                MessageBox.Show("Ошибка выгрузки документа в Excel: " + ex, "Ошибка", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void OpenFileOrDirectory(string filename)
        {
            var dialogResult = MessageBox.Show("Показать результаты?", "Информация", MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);

            if (dialogResult == DialogResult.Yes)
            {
                Process.Start(filename);
            }
        }

        private void BtnUpload_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_currentReport))
            {
                return;
            }

            try
            {
                if (SuccessStatuses.Contains(_processor.Report.Status) || _processor.Report.Status == ReportStatus.New)
                {
                    throw new Exception("Скан можно закачать только для отчетов, которые находится в статусах: " +
                        "'Сохранен в БД', 'Загружен скан', 'Отправлен на доработку'");
                }

                openFileDialog1.Filter = "PDF | *.pdf";
                if (openFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                string filename = openFileDialog1.FileName;
                var extension = Path.GetExtension(filename)?.ToLower() ?? "";

                if (extension != ".pdf")
                {
                    throw new Exception("Можно загрузить только документы с расширением PDF");
                }

                string savedFileName = GetFileName(extension);
                _ftpProcessor.UploadFileToWs(filename, _processor.FilialCode, savedFileName, _client);

                _processor.SaveScan(savedFileName, 1);
                _processor.Report.Status = ReportStatus.Scan;
                ReportTree.SelectedNode.BackColor = ColorScan;

                MessageBox.Show("Файл успешно загружен на сервер", "Загрузка завершена", MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log.Error(ex, $"Error saving scan of file");
                MessageBox.Show("Ошибка сохранения скана: " + ex.Message, "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void CollectSummaryReport()
        {
            var filialList = new List<string>();
            if (!string.IsNullOrEmpty(TxtbFilials.Text))
            {
                var parsedFilialList = TxtbFilials.Text.Replace(", ", ",").Split(',');
                filialList.AddRange(parsedFilialList.Select(f => _regions.Single(x => x.Value == f).Key));
            }

            var yymmStart = YymmUtils.GetYymmFromInt(NumStart.Value, CmbStart.SelectedValue);
            var yymmEnd = YymmUtils.GetYymmFromInt(NumEnd.Value, CmbEnd.SelectedValue);

            ReportStatus status = Enum.TryParse(CmbFilterType.SelectedValue.ToString(), out ReportStatus enumStatus)
                ? enumStatus
                : ReportStatus.Saved;

            DataSource datasource = Enum.TryParse(CmbFilterType.SelectedValue.ToString(), out DataSource enumdatasource)
                ? enumdatasource
                : DataSource.New;

            try
            {
                _processor.FindReports(filialList, yymmStart, yymmEnd, status, datasource);
                _processor.FillDataGridView(_processor.GetCurrentTheme());
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error getting summary report by filter");
                MessageBox.Show("Ошибка получения сумммарного отчета: " + ex.Message, "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void CheckUpdateApplication(bool idApplicationStart)
        {
            var updater = new ApplicationUpdater(_ftpProcessor, _client);
            updater.GetDll();
            updater.UpdateApp(idApplicationStart);
        }

        private bool CheckComment() =>
            _processor.Report.IdFlow != 0 && _client.IsReportHasComments(_processor.Report.IdFlow);

        private void ChangeIndexComboBox(DataGridView dgw, ComboBox cmb, TextBox txtb)
        {
            if (_processor == null)
            {
                return;
            }
            if (_processor.ThemesList.Select(x => x.Key).Contains(cmb.Text) || cmb.Text == "Свод")
            {
                _processor.MapReportFromDgv(_processor.OldTheme);
                _processor.OldTheme = cmb.Text;
                _processor.CreateReportForm(cmb.Text);
                _processor.FillDataGridView(cmb.Text);
                //_processor.SetTotalColumn();


                txtb.Text = cmb.SelectedValue.ToString();
            }

            if (!CurrentUser.IsMain)
            {
                BtnUploaded.Visible = _processor.IsVisibleBtnDownloadExcel();
                BtnHandle.Visible = _processor.IsVisibleBtnHandle();
                dgw.AllowUserToDeleteRows = _processor.IsVisibleBtnDownloadExcel();
                separatorExcel.Visible = _processor.IsVisibleBtnDownloadExcel();
            }
            else
            {
                BtnUploaded.Visible = false;
                BtnHandle.Visible = false;
            }
        }



        private string GetFileName(string extension) =>
            $"{_processor.FilialCode}_{_processor.SmallName}_{_processor.Report.Yymm}{extension}";

        private void OpenConsolidateReportForm(ConsolidateReport consolidateReport)
        {
            using var form = new ConsolidateForm(_client, _regions, consolidateReport, _processor.FilialName);
            form.ShowDialog();
        }


        private void CollectReportDataUploaded()
        {
            if (string.IsNullOrEmpty(_currentReport))
            {
                return;
            }
            _processor.Report.DataSource = DataSource.Excel;
            openFileDialog1.Filter = "Excel | *.xls; *.xlsx";
            string theme = _processor.GetCurrentTheme();
            if (openFileDialog1.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            try
            {
                var excelCollector = _excelCollectorFactory.GetExcelCollector(_currentReport);
                excelCollector.Collect(openFileDialog1.FileName, theme, _processor.Report);

                _processor.CreateReportForm(theme);
                _processor.FillDataGridView(theme);
                MessageBox.Show("Данные успешно загружены на форму", "Загрузка");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error collecting report data from excel");
                MessageBox.Show("Ошибка получения данных из Excel " + ex.Message, "Ошибка", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void CollectDynamicReportDataUploaded()
        {

            var tag = ReportTree.SelectedNode.Tag as ReportNodeTag;
            if (tag.IdReport == 0)
                return;

            openFileDialog1.Filter = "Excel | *.xls; *.xlsx";
            if (openFileDialog1.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            try
            {
                var collector = new DynamicReportExcelCollector();

                collector.Collect(openFileDialog1.FileName, _dynamicReportProcessor, _dynamicReportProcessor.Report);
                _dynamicReportProcessor.FillThemeData(DgvQuery);


                MessageBox.Show("Данные успешно загружены на форму", "Загрузка");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error collecting report data from excel");
                MessageBox.Show("Ошибка получения данных из Excel " + ex.Message, "Ошибка", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void RefuseReport()
        {
            if (string.IsNullOrEmpty(_currentReport))
            {
                return;
            }

            try
            {
                if (!SuccessStatuses.Contains(_processor.Report.Status))
                {
                    throw new Exception("Вернуть на доработку можно только отчеты в статусах: " +
                        "'Направлен в ЦО' или 'Утвержден'");
                }

                _processor.ChangeStatus(ReportStatus.Refuse);

                ReportTree.SelectedNode.BackColor = ColorRefuse;
                _processor.Report.Status = ReportStatus.Refuse;
                BtnSubmit.Enabled = true;

                MessageBox.Show("Отчет отправлен на доработку!", "Отправка на доработку", MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error sending report to refuse");
                MessageBox.Show("Ошибка отправки отчета на доработку!" + ex.Message, "Ошибка", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }



        private void AutoFillReportFromPrevious()
        {
            if (string.IsNullOrEmpty(_currentReport))
            {
                return;
            }

            try
            {
                var yymmPrev = YymmUtils.ConvertYymmToDate(_processor.Report.Yymm).AddMonths(-1).ToString("yyMM");
                var report = _processor.CollectReportFromWs(yymmPrev);

                if (report != null)
                {
                    _processor.MapForAutoFill(report);
                    _processor.FillDataGridView(_processor.GetCurrentTheme());
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error autofill report");
                MessageBox.Show("Ошибка автозаполнения отчета!" + ex.Message, "Ошибка", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void SubmitReport()
        {
            if (_processor.Report.Status == ReportStatus.Refuse)
            {
                MessageBox.Show(
                    "Данный отчет был возвращен на доработку. Для повторной сдачи отчета необходимо перезакачать скан",
                    "Предупреждение!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (_processor.Report.Status != ReportStatus.Scan)
            {
                MessageBox.Show("Можно сдавать только те отчеты, у которых загружен скан.", "Предупреждение!",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var dialogResult = MessageBox.Show(
                "Вы уверены что хотите сдать отчет? Дальнейшее редактирование данной версии будет невозможно",
                "Информация!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (dialogResult != DialogResult.Yes)
            {
                return;
            }

            try
            {
                _processor.ChangeStatus(ReportStatus.Submit);
                ReportTree.SelectedNode.BackColor = ColorSubmit;
                _processor.Report.Status = ReportStatus.Submit;
                BtnSubmit.Enabled = false;

                SetReportInterface();
                MessageBox.Show("Отчет сдан!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error submiting report");
                MessageBox.Show("Ошибка сдачи отчета:" + ex.Message, "Ошибка!", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void DoneReport()
        {
            if (_processor.Report.Status != ReportStatus.Submit)
            {
                MessageBox.Show("Можно утверждать только отчеты в статусе 'Отчет направлен в ЦО'", "Ошибка!",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                _processor.ChangeStatus(ReportStatus.Done);
                ReportTree.SelectedNode.BackColor = ColorIsDone;
                _processor.Report.Status = ReportStatus.Done;
                BtnSubmit.Enabled = false;

                MessageBox.Show("Отчет утвержден!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                MessageBox.Show("Ошибка сохранения отчета!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ChkbFilter_CheckedChanged(object sender, EventArgs e)
        {
            if (_currentReportName != null)
            {
                PanelFilterInt.Enabled = ChkbFilter.Checked;
                BtnOpenScan.Enabled = !ChkbFilter.Checked;
                BtnRefuse.Enabled = !ChkbFilter.Checked;
                BtnSubmit.Enabled = !ChkbFilter.Checked;

                TbControl.SelectedTab.Text = ChkbFilter.Checked
                    ? $"{_currentReportName} | Фильтр"
                    : $"{_currentReportName} | {_yymm} | {_regions.Single(x => x.Key == _processor.FilialCode).Value}";
            }
        }

        private void TabReport_DrawItem(object sender, DrawItemEventArgs e)
        {
            var g = e.Graphics;
            var tp = TbControl.TabPages[e.Index];

            var sf = new StringFormat { Alignment = StringAlignment.Center };
            var headerRect = new RectangleF(e.Bounds.X, e.Bounds.Y, e.Bounds.Width, e.Bounds.Height);
            var sb = new SolidBrush(SystemColors.Control);

            Color color = _processor.ColorReport;

            if (_isQuery)
            {
                if (ReportTree.SelectedNode != null)
                {
                    var tag = ReportTree.SelectedNode.Tag as ReportNodeTag;
                    if (tag != null)
                    {
                        if (tag.idFlow != 0)
                        {
                            var flow = _client.GetReportDynamicFlowById(tag.idFlow);
                            color = (Color)GetColorForNode(flow.Status);

                        }
                    }
                }

            }
            if (TbControl.SelectedIndex == e.Index)
            {
                sb.Color = color;
            }

            g.FillRectangle(sb, e.Bounds);
            g.DrawString(tp.Text, TbControl.Font, new SolidBrush(Color.Black), headerRect, sf);
        }

        private void TabReport_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (TbControl.SelectedTab == null)
            {
                return;
            }

            _currentReport = _tabControlMap[TbControl.SelectedTab];
            _processor = _processorMap[_currentReport];
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_currentReport))
            {
                return;
            }

            try
            {
                SerializeReport();
                MessageBox.Show("Успешно сохранено", "Ок!", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка сериализации");
                MessageBox.Show($"Ошибка сериализации: {ex}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DoneDynamicReportReport()
        {

            if (_currentDynamicReportFlow == 0)
            {
                return;
            }
            var flow = _dynamicReportProcessor.GetReportDynamicFlow(_currentDynamicReportFlow);
            if (flow.Status != ReportStatus.Submit)
            {
                MessageBox.Show("Можно утверждать только отчеты в статусе 'Отчет направлен в ЦО'", "Ошибка!",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _client.ChangeDynamicReportStatus(_currentDynamicReportFlow, ReportStatus.Done);
            ReportTree.SelectedNode.BackColor = ColorIsDone;

            MessageBox.Show("Отчет утвержден!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);


        }

        private void SubmitDynamicReport()
        {
            var tag = ReportTree.SelectedNode.Tag as ReportNodeTag;
            if (tag.idFlow == 0)
            {
                return;
            }
            var flow = _dynamicReportProcessor.GetReportDynamicFlow(tag.idFlow);

            if (flow.Status == ReportStatus.Submit)
            {
                return;
            }

            var dialogResult = MessageBox.Show(
          "Вы уверены что хотите сдать отчет? Дальнейшее редактирование данной версии будет невозможно",
          "Информация!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (dialogResult == DialogResult.Yes)
            {
                _client.ChangeDynamicReportStatus(tag.idFlow, ReportStatus.Submit);
                ReportTree.SelectedNode.BackColor = ColorSubmit;
                MessageBox.Show("Отчет сдан!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


        }


        private Color? GetColorForNode(ReportStatus? status)
        {
            switch (status)
            {
                case ReportStatus.Done:
                    return GlobalConst.ColorIsDone;
                case ReportStatus.Refuse:
                    return GlobalConst.ColorRefuse;
                case ReportStatus.Submit:
                    return GlobalConst.ColorSubmit;
                case ReportStatus.Scan:
                    return GlobalConst.ColorScan;
                case ReportStatus.Saved:
                    return GlobalConst.ColorBd;
                default:
                    return null;
            }
        }

        private void BtnSend_Click(object sender, EventArgs e)
        {
            if (_isQuery)
            {
                if (CurrentUser.IsMain)
                {
                    DoneDynamicReportReport();
                }
                else
                {
                    SubmitDynamicReport();
                }

                return;
            }
            if (string.IsNullOrEmpty(_currentReport))
            {
                return;
            }

            if (CurrentUser.IsMain)
            {
                DoneReport();
            }
            else
            {
                SubmitReport();
            }
        }

        private void BtnCommentReport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_currentReport))
            {
                return;
            }

            using var form = new CommentForm(_client, _processor.Report);
            form.ShowDialog();
        }

        private void MenuSetting_Click(object sender, EventArgs e)
        {
            using var form = new SettingsForm(_client);
            form.ShowDialog();
        }

        private void РассылкаУведомленийToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!_isQuery)
            {
                using var form = new NotificationForm(_client, _regions, _reportsDictionary);
                form.ShowDialog();
            }
            else
            {
                using var form = new NotificationForm(_client, _dynamicReportProcessor.Report, _dynamicReportProcessor.Report.Executors);
                form.ShowDialog();
            }

        }

        private void BtnAutoFill_Click(object sender, EventArgs e) =>
            AutoFillReportFromPrevious();

        private void BtnUploaded_Click(object sender, EventArgs e)
        {
            if (!_isQuery)
            {
                CollectReportDataUploaded();
                SaveReportToDB_DataSourceToExcel();
            }
            else
            {
                CollectDynamicReportDataUploaded();

            }


        }

        private void BtnSummary_Click(object sender, EventArgs e)
        {
            if (!_isQuery)
            {
                CollectReportDataUploaded();
                SaveReportToDB_DataSourceToExcel();
            }
            else
            {
                CollectDynamicReportDataUploaded();

            }


        }

        private void BtnHandle_Click(object sender, EventArgs e)
        {
            if (!_isQuery)
            {
                SaveReportToDB_DataSourceToHandle();
            }
        }

        private void ReportTree_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (_isQuery)
            {
                CreateDynamicReportTab();
            }
            else
            {
                CreateReportTab();
            }



        }


        private void BtnRefresh_Click(object sender, EventArgs e) =>
            CreateTreeView();

        private void BtnRefuse_Click(object sender, EventArgs e)
        {
            if (_isQuery)
            {
                RefuseDynamicReport();
            }
            else
            {
                RefuseReport();
            }


        }

        private void RefuseDynamicReport()
        {

            if (_currentDynamicReportFlow == 0)
            {
                return;
            }
            var flow = _dynamicReportProcessor.GetReportDynamicFlow(_currentDynamicReportFlow);
            if (!(flow.Status == ReportStatus.Done || flow.Status == ReportStatus.Submit))
            {
                MessageBox.Show("Вернуть на доработку можно только отчеты в статусах: " +
                        "'Направлен в ЦО' или 'Утвержден'", "Ошибка!",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _client.ChangeDynamicReportStatus(_currentDynamicReportFlow, ReportStatus.Refuse);
            ReportTree.SelectedNode.BackColor = ColorRefuse;
            MessageBox.Show("Отчет отправлен на доработку!", "Отправка на доработку", MessageBoxButtons.OK,
                 MessageBoxIcon.Information);
        }

        private void TbtnOpenScan_Click(object sender, EventArgs e)
        {

            if (!_isQuery)
            {
                using var form = new ScanForm(_processor, ReportTree, _client);
                form.ShowDialog();
            }
            else
            {
                var reportTag = ReportTree.SelectedNode.Tag as ReportNodeTag;
                if (reportTag != null)
                {
                    if (reportTag.idFlow == 0)
                    {
                        if (!CurrentUser.IsMain)
                            MessageBox.Show("Сначала необходимо сохранить отчёт!", "Информация", MessageBoxButtons.OK,
                   MessageBoxIcon.Information);
                        return;
                    }

                    using var form = new ScanDynamicForm(_dynamicReportProcessor, reportTag, ReportTree, _client);
                    form.ShowDialog();
                }
            }
        }

        private void AddValueInTextBox()
        {
            bool all = false;
            if (CmbFilials.SelectedIndex == _regions.Count)
            {
                all = true;
            }
            GlobalUtils.AddValueInTextBox(CmbFilials, TxtbFilials, true, all);

        }

        private void Open_ReleaseChangelogForm()
        {
            using var releaseChangelogForm = new ReleaseChangelogForm();
            releaseChangelogForm.ShowDialog();
        }

        private void MenuChangelog_Click(object sender, EventArgs e) =>
            Open_ReleaseChangelogForm();






        private void MenuCheckPoUpdate_Click(object sender, EventArgs e) =>
            CheckUpdateApplication(false);

        private void BtnCreateExcel_Click(object sender, EventArgs e)
        {
            if (_isQuery)
            {
                UploadToExcelDynamicReport();
            }
            else
            {
                UploadToExcel();
            }
        }

        private void MenuExit_Click(object sender, EventArgs e) =>
            Close();

        private void CmbThemes_SelectedIndexChanged(object sender, EventArgs e) =>
            ChangeIndexComboBox(DgwReportIizl, CmbIizl, TxtbIizl);

        private void Cmb262_SelectedIndexChanged(object sender, EventArgs e) =>
            ChangeIndexComboBox(DgwReport262, Cmb262, Txtb262);

        private void Cmb294_SelectedIndexChanged(object sender, EventArgs e) =>
            ChangeIndexComboBox(DgwReport294, Cmb294, Txtb294);

        private void CmbPg_SelectedIndexChanged(object sender, EventArgs e) =>
            ChangeIndexComboBox(DgwReportPgQ, CmbPgQ, TxtbPgQ);

        private void CmbPg_SelectedIndexChanged_1(object sender, EventArgs e) =>
            ChangeIndexComboBox(DgwReportPg, CmbPg, TxtbPg);

        private void CmbZpz_SelectedIndexChanged(object sender, EventArgs e) =>
        ChangeIndexComboBox(DgwReportZpzQ, CmbZpzQ, TxtbZpzQ);

        private void CmbZpz2025_SelectedIndexChanged(object sender, EventArgs e) =>
        ChangeIndexComboBox(DgvReportZpzQ2025, CmbZpzQ2025, TxtbZpzQ2025);

        private void CmbCadre_SelectedIndexChanged(object sender, EventArgs e) =>
        ChangeIndexComboBox(DgvCadre, CmbCadre, TxtbCadre);

        private void CmbDoff_SelectedIndexChanged(object sender, EventArgs e) =>
        ChangeIndexComboBox(DgvDoff, CmbDoff, TbDoff);

        private void CmbReqVCR_SelectedIndexChanged(object sender, EventArgs e) =>
        ChangeIndexComboBox(DgvReqVCR, CmbReqVCR, TxtbReqVCR);

        private void CmbZpzLethal_SelectedIndexChanged(object sender, EventArgs e) =>
        ChangeIndexComboBox(DgwReportZpzLethal, CmbZpzLethal, TxtbZpzLethal);

        private void CmbZpz2025Lethal_SelectedIndexChanged(object sender, EventArgs e) =>
ChangeIndexComboBox(DgvReportZpzLethal2025, CmbZpzLethal2025, TxtbZpzLethal2025);

        private void CmbDoff_SelectedIndexChanged_1(object sender, EventArgs e) =>
        ChangeIndexComboBox(DgvDoff, CmbDoff, TbDoff);

        private void CmbViolMEE_SelectedIndexChanged_1(object sender, EventArgs e) =>
ChangeIndexComboBox(DgvViolMEE, CmbViolMEE, TbViolMEE);

        private void CmbViolEKMP_SelectedIndexChanged_1(object sender, EventArgs e) =>
ChangeIndexComboBox(DgvViolEKMP, CmbViolEKMP, TbViolEKMP);


        private void CmbVerifyPlan_SelectedIndexChanged_1(object sender, EventArgs e) =>
ChangeIndexComboBox(DgvVerifyPlan, CmbVerifyPlan, TbVerifyPlan);

        private void CmbMonthlyVol_SelectedIndexChanged_1(object sender, EventArgs e) =>
ChangeIndexComboBox(DgvMonthlyVol, CmbMonthlyVol, TbMonthlyVol);

        private void CmbZpz2025_SelectedIndexChanged_1(object sender, EventArgs e) =>
            ChangeIndexComboBox(DgvReportZpz2025, CmbZpz2025, TxtbZpz2025);

        private void CmbZpz10_2025_SelectedIndexChanged(object sender, EventArgs e) =>
        ChangeIndexComboBox(DgvReportZpz10_2025, CmbZpz10_2025, TxtbZpz10_2025);

        private void CmbZpz_SelectedIndexChanged_1(object sender, EventArgs e) =>
    ChangeIndexComboBox(DgwReportZpz, CmbZpz, TxtbZpz);

        private void CmbZpz10_SelectedIndexChanged(object sender, EventArgs e) =>
        ChangeIndexComboBox(DgwReportZpz10, CmbZpz10, TxtbZpz10);

        private void BtnPlus_Click(object sender, EventArgs e) =>
             AddValueInTextBox();

        private void BtnMinus_Click(object sender, EventArgs e) =>
            GlobalUtils.DeleteValueFromTextBox(TxtbFilials);

        private void BtnClearReport_Click(object sender, EventArgs e) =>
            CreateNewFilter();

        private void BtnFindReport_Click(object sender, EventArgs e) =>
            CollectSummaryReport();

        private void BtnOpen_Click(object sender, EventArgs e) =>
            DeserializeReport();

        private void BtnSaveToDb_Click(object sender, EventArgs e)
        {
            if (_isQuery)
            {
                SaveDynamycReportToDb();
            }
            else
            {
                SaveReportToDb();
            }
        }


        private void отделЗПЗИЭКМПToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.ConsolidateCadreT1);

        private void оИИЗПЗToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.ConsolidateCadreT2);

        private void СводКТаблице1ToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.Consolidate262T1);

        private void СводКТаблице2ToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.Consolidate262T2);

        private void СводКТаблице3ToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.Consolidate262T3);

        private void КонтрольЗПЗToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.ControlZpzMonthly);

        private void КонтрольЗПЗ2025ToolStripMenuItem_Click(object sender, EventArgs e) =>
        OpenConsolidateReportForm(ConsolidateReport.ControlZpz2025Quarterly);

        private void СуммарныйОтчетПоФилиалуToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.ConsolidateFilial294);

        private void ИтоговыйОтчетПоВсемФилиаламToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.ConsolidateFull294);

        private void ОтчетДляСайтаToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.ZpzWebSite);

        private void ОтчетДляСайта2023ToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.ZpzWebSite2023);

        private void ОтчетДляСайта2025ToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.ZpzWebSite2025);

        private void ОтчетРезультатыМЭКToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.ZpzTable5);

        private void ОтчетНарушенияПоОбращениямЗЛToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.ViolationsOfAppeals);

        private void ОбъёмыЕжемесячныеToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.FFOMSMonthlyVol);

        private void НарушенияМЭЭToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.FFOMSViolMEE);

        private void НарушенияЭКМПToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.FFOMSViolEKMP);

        private void ПланыПроверокToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.FFOMSVerifyPlan);

        private void ВнеплановыеЭкспертизыToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.FFOMSTargetedExp);

        private void ОнкологияХТToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.FFOMSOncoCT);

        private void ФФОМСКадрыToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.FFOMSPersonnel);

        private void ФФОМСЛетальныеЭКМПToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.FFOMSLethalEKMP);

        private void ФФОМСОбъемыПоВидамПомощиToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.FFOMSVolumesByTypes);

        private void КонтрольЗПЗежемесячнаяToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.ControlZpzQuarterly);

        //private void КонтрольЗПЗ2025ежемесячнаяToolStripMenuItem_Click(object sender, EventArgs e) =>
        //    OpenConsolidateReportForm(ConsolidateReport.ControlZpz2025Quarterly);

        private void КонтрольЗПЗ2023ToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.ControlZpz2023Quarterly);

        private void КонтрольЗПЗ2023FullToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.ControlZpz2023FullQuarterly);

        private void КонтрольЗПЗ2023SingleToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.ControlZpz2023SingleQuarterly);

        private void КонтрольЗПЗ2025FullToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.ControlZpz2025FullQuarterly);

        private void КонтрольЗПЗ2025SingleToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.ControlZpz2025SingleQuarterly);

        private void онкологияToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.Onko);

        private void онкологияквартальныйToolStripMenuItem_Click(object sender, EventArgs e) =>
            OpenConsolidateReportForm(ConsolidateReport.OnkoQuarterly);

        //private void ИсполненииЦПНПToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    OpenConsolidateReportForm(ConsolidateReport.CnpnQuarterly);
        //}

        //private void ЕжемесячныйToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    OpenConsolidateReportForm(ConsolidateReport.CnpnMonthly);
        //}

        private void СердечнососудистыеЗаболеванияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.Cardio);
        }

        private void ЦПНПToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void ЦПНПежемесячныйToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.CnpnMonthly);
        }

        private void ЦПНПквартальныйToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.CnpnQuarterly);
        }

        private void вновьЗСИУДToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.ConsQuantityAR);
        }

        private void ЧислСводКварталToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.ConsQuantityQ);
        }

        private void выполнениеПланаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.ConsQuantityFP);
        }

        private void сведенияЧисленностьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.ConsQuantityInformation);
        }


        private void диспанцеризацияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.Disp);

        }

        private void CmbFilterType_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void CmbFilials_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void serviceMenu_Click(object sender, EventArgs e)
        {

        }

        private void форма294ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void создатьОтчётнуюФормуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var constructorForm = new ConstuctorForm(_client);
            constructorForm.ShowDialog();

        }

        private void CmbTypeTree_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void CmbTypetTree_SelectedIndexChanged(object sender, EventArgs e)
        {
            int SelecetedYear = (int)TreeYear.Value;
            switch (CmbTypeTree.SelectedValue)
            {
                case "Отчёты":
                    TbControl.TabPages.Remove(PageQuery);
                    CreateTreeView();
                    _isQuery = false;
                    ChkbFilter.Enabled = true;
                    break;
                case "Запросы":
                    _isQuery = true;
                    ChkbFilter.Enabled = false;
                    ReportTree.Nodes.Clear();
                    _reportView.CreateTreeViewQuery(SelecetedYear);
                    if (!CurrentUser.IsMain)
                        BtnUploaded.Visible = false;
                    BtnHandle.Visible = false;
                    TbControl.TabPages.Remove(Page262);
                    TbControl.TabPages.Remove(Page294);
                    TbControl.TabPages.Remove(PageIizl);
                    TbControl.TabPages.Remove(PagePg);
                    TbControl.TabPages.Remove(PagePgQ);
                    TbControl.TabPages.Remove(PageZpz);
                    TbControl.TabPages.Remove(PageZpz10);
                    TbControl.TabPages.Remove(PageZpzQ);
                    TbControl.TabPages.Remove(PageZpz2025);
                    TbControl.TabPages.Remove(PageZpz10_2025);
                    TbControl.TabPages.Remove(PageZpzQ2025);
                    TbControl.TabPages.Remove(PageQuery);
                    TbControl.TabPages.Remove(PageCadre);
                    TbControl.TabPages.Remove(PageReqVCR);
                    TbControl.TabPages.Remove(PageOtclkInfrorm);
                    TbControl.TabPages.Remove(PageOped);
                    TbControl.TabPages.Remove(PageOpedU);
                    TbControl.TabPages.Remove(PageOpedQ);
                    TbControl.TabPages.Remove(PageProposal);
                    TbControl.TabPages.Remove(tpOpedFinance);
                    TbControl.TabPages.Remove(tpOpedFinance3);
                    TbControl.TabPages.Remove(tpIizl2022);
                    TbControl.TabPages.Remove(PageFssMonitoring);
                    TbControl.TabPages.Remove(PageMonitoringVCR);
                    TbControl.TabPages.Remove(PageQuantity);
                    TbControl.TabPages.Remove(PageTarAllow);
                    TbControl.TabPages.Remove(PageDoff);
                    TbControl.TabPages.Remove(PageViolMEE);
                    TbControl.TabPages.Remove(PageViolEKMP);
                    TbControl.TabPages.Remove(PageVerifyPlan);
                    break;
            }
        }

        private void CmbQuery_SelectedIndexChanged(object sender, EventArgs e)
        {
            _dynamicReportProcessor.oldPageIndex = _dynamicReportProcessor._pageIndex;
            _dynamicReportProcessor._pageIndex = CmbQuery.SelectedIndex;
            _dynamicReportProcessor.SetData(DgvQuery, _dynamicReportProcessor.oldPageIndex);
            _dynamicReportProcessor.SetDgv(DgvQuery, CmbQuery.Text);
            _dynamicReportProcessor.FillThemeData(DgvQuery);
            TbxQuery.Text = _dynamicReportProcessor.GetDescriptionPage(CmbQuery.Text);

            if (ReportTree.SelectedNode != null)
            {
                if (CurrentUser.IsMain)
                {
                    if (ReportTree.SelectedNode.Level != 1)
                    {
                        PageQuery.Text = ReportTree.SelectedNode.Parent.Parent.Text + "\n" + ReportTree.SelectedNode.Text + "\n" + ReportTree.SelectedNode.Parent.Text;

                    }
                    else
                    {
                        PageQuery.Text = ReportTree.SelectedNode.Parent.Text + "\n" + ReportTree.SelectedNode.Text;

                    }

                }
                else
                {
                    if (ReportTree.SelectedNode.Parent != null)
                        PageQuery.Text = ReportTree.SelectedNode.Parent.Text + "\n" + ReportTree.SelectedNode.Text + "\n";

                }

            }


            if (_dynamicReportProcessor.Report.Id == 33 /*&& ReportTree.SelectedNode.Level == 2*/) // todo убрать. для проверки 2022
            {

                _dynamicReportProcessor.TuneProverkaTfomsTables(CmbQuery.Text.Substring(CmbQuery.Text.Length - 4, 4),
                   CurrentUser.IsMain && ReportTree.SelectedNode.Level == 2 ? CurrentUser.Regions.FirstOrDefault(x => x.Value == ReportTree.SelectedNode.Text).Key : CurrentUser.FilialCode);
                //_dynamicReportProcessor.SetFFOMSCheck2022LetalData(CmbQuery.Text.Substring(CmbQuery.Text.Length - 4, 4), CurrentUser.Regions.FirstOrDefault(x => x.Value == ReportTree.SelectedNode.Text).Key);
            }


        }

        private void редактироватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (ReportTree.SelectedNode == null)
            {
                return;
            }
            if (ReportTree.SelectedNode.Level == 0 || ReportTree.SelectedNode.Level == 2)
            {
                return;
            }
            var selecteNode = ReportTree.SelectedNode;
            for (int i = 0; i < selecteNode.Nodes.Count; i++)
            {
                if (selecteNode.Nodes[i].BackColor != Color.Empty)
                {
                    MessageBox.Show("Редактирование невозможно! Некоторые филиалы уже заполнили данную отчётную форму.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            var selectedReport = ReportTree.SelectedNode.Tag as ReportNodeTag;

            if (selectedReport.IdReport != 0)
            {
                using var form = new ConstuctorForm(_client, selectedReport.IdReport);
                form.ShowDialog();
            }


        }

        private void DgvQuery_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int rowIndex = DgvQuery.CurrentRow.Index + 1;
            int colIndex = DgvQuery.CurrentCell.ColumnIndex + 1;

            if (DgvQuery.Columns[0].Name == "Наименование показателя")
            {
                colIndex -= 2;
            }

            var page = _dynamicReportProcessor.Report.Page.ElementAt(CmbQuery.SelectedIndex).Value;


            try
            {
                var currentColumn = page.Columns.Where(x => !x.IsGroup).FirstOrDefault(x => Convert.ToInt32(x.Index) == colIndex);
                if (currentColumn == null)
                {
                    foreach (var item in page.Columns.Where(x => x.IsGroup))
                    {
                        if (item.Columns.FirstOrDefault(x => Convert.ToInt32(x.Index) == colIndex) != null)
                        {
                            currentColumn = item.Columns.FirstOrDefault(x => Convert.ToInt32(x.Index) == colIndex);

                        }
                    }
                }

                var currentRow = page.Rows.FirstOrDefault(x => Convert.ToInt32(x.Index) == rowIndex);
                string message = string.Empty;

                if (currentColumn != null)
                {
                    if (!string.IsNullOrEmpty(currentColumn.Description))
                    {
                        message += String.Format($"Столбец:{currentColumn.Description.Trim()}") + Environment.NewLine;

                    }
                }


                if (currentRow != null)
                {
                    if (!string.IsNullOrEmpty(currentRow.Description))
                    {
                        message += String.Format($"Строка:{currentRow.Description.Trim()}");

                    }
                }

                TbxEmentInfo.Text = message;

            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error submiting report");
                //MessageBox.Show("Ошибка сдачи отчета:" + ex.Message, "Ошибка!", MessageBoxButtons.OK,
                //    MessageBoxIcon.Error);
            }



        }

        private void Con_Click(object sender, EventArgs e)
        {

        }

        private void летальныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.Letal);
        }

        private void летальные2023ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.Letal2023);
        }

        private void справочникиToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void электронныеАдресаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var form = new EmailForm(_client);
            form.ShowDialog();
        }

        private void отправитьToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void DgvReportOped_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DgvReportOpedU_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DgvReportOped_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {


        }

        private void DgvReportOpedU_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {


        }

        private void DgvReportOped_KeyDown(object sender, KeyEventArgs e)
        {
            _processor.CallculateCells();
        }

        private void DgvReportOpedU_KeyDown(object sender, KeyEventArgs e)
        {
            _processor.CallculateCells();
        }

        private void DgvReportOped_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void DgvReportOpedU_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void DgvReportZpz10_KeyPress(object sender, KeyPressEventArgs e)
        {
            (_processor as ReportZpz10Processor).SetFormula();
        }

        private void DgvReportZpz10_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            (_processor as ReportZpz10Processor).SetFormula();
        }


        private void DgvReportZpz10_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportZpz10Processor).SetFormula();
        }

        private void DgvReportZpzQ_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportZpzQProcessor).SetFormula();
        }

        private void DgvReportZpzLethal_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportZpzLethalProcessor).SetFormula();
        }


        private void DgvReportZpz10_2025_KeyPress(object sender, KeyPressEventArgs e)
        {
            (_processor as ReportZpz10_2025Processor).SetFormula();
        }

        private void DgvReportZpz10_2025_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            (_processor as ReportZpz10_2025Processor).SetFormula();
        }


        private void DgvReportZpz10_2025_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportZpz10_2025Processor).SetFormula();
        }

        private void DgvReportZpzQ2025_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportZpzQ2025Processor).SetFormula();
        }

        private void DgvReportZpz2025Lethal_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportZpz2025LethalProcessor).SetFormula();
        }

        private void DgvReportMVCR_KeyPress(object sender, KeyPressEventArgs e)
        {
            (_processor as MonitoringVCRProcessor).SetFormula();
        }

        private void DgvReportMVCR_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            (_processor as MonitoringVCRProcessor).SetFormula();
        }


        private void DgvReportMVCR_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as MonitoringVCRProcessor).SetFormula();
        }

        private void DgvReportZpz10_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportZpz10Processor).SetFormula();
        }

        private void DgvReportZpz10_2025_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportZpz10_2025Processor).SetFormula();
        }

        private void DgvReportOped_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            _processor.CallculateCells();
        }

        private void DgvReportOpedU_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            _processor.CallculateCells();
        }


        private void DgvReportViolMEE_KeyPress(object sender, KeyPressEventArgs e)
        {
            (_processor as ReportViolMEEProcessor).SetFormula();
        }

        private void DgvReportViolMEE_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            (_processor as ReportViolMEEProcessor).SetFormula();
        }


        private void DgvReportViolMEE_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportViolMEEProcessor).SetFormula();
        }


        private void DgvReportViolEKMP_KeyPress(object sender, KeyPressEventArgs e)
        {
            (_processor as ReportViolEKMPProcessor).SetFormula();
        }

        private void DgvReportViolEKMP_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            (_processor as ReportViolEKMPProcessor).SetFormula();
        }


        private void DgvReportViolEKMP_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportViolEKMPProcessor).SetFormula();
        }





        private void DgvReportVerifyPlan_KeyPress(object sender, KeyPressEventArgs e)
        {
            (_processor as ReportVerifyPlanProcessor).SetFormula();
        }

        private void DgvReportMonthlyVol_KeyPress(object sender, KeyPressEventArgs e)
        {
            (_processor as ReportMonthlyVolProcessor).SetFormula();
            (_processor as ReportMonthlyVolProcessor).CalculateCells();
        }

        private void DgvReportMonthlyVol_Move(object sender, MouseEventArgs e)
        {
            (_processor as ReportMonthlyVolProcessor).SetFormula();
            (_processor as ReportMonthlyVolProcessor).CalculateCells();
        }

        private void DgvReportVerifyPlan_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            (_processor as ReportVerifyPlanProcessor).SetFormula();
        }


        private void DgvReportVerifyPlan_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportVerifyPlanProcessor).SetFormula();
        }

        private void DgvReportMonthlyVol_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            (_processor as ReportMonthlyVolProcessor).SetFormula();
        }


        private void DgvReportMonthlyVol_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportMonthlyVolProcessor).SetFormula();
        }


        private void сводToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.ConsolidateOped);
        }

        private void сводUToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.ConsolidateOpedUnplanned);
        }

        private void CbxOped_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangeIndexComboBox(DgvReportOped, CbxOped, TxtbOped);
        }
        private void CbxOpedU_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangeIndexComboBox(DgvReportOpedU, CbxOpedU, TxtbOpedU);
        }

        private void CbxOtclkInfrorm_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangeIndexComboBox(DgvOtclkInfrorm, CbxOtclkInfrorm, TxtOtclkInfrorm);
        }

        private void CmbPageCadre_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangeIndexComboBox(DgvCadre, CmbCadre, TxtbCadre);
        }

        private void CmbPageReqVCR_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangeIndexComboBox(DgvReqVCR, CmbReqVCR, TxtbReqVCR);
        }

        private void DgvOtclkInfrorm_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DgvCadre_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DgvReqVCR_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void DgvOtclkInfrorm_KeyPress(object sender, KeyPressEventArgs e)
        {
            (_processor as ReportInfrormationResponseProcessor).SetFormula();
        }

        private void DgvCadre_KeyPress(object sender, KeyPressEventArgs e)
        {
            (_processor as ReportCadreProcessor).SetFormula();
            (_processor as ReportCadreProcessor).SetFormula();
        }

        private void DgvReqVCR_KeyPress(object sender, KeyPressEventArgs e)
        {
            (_processor as ReportReqVCRProcessor).SetFormula();
        }

        private void DgvEffectiveness_keyPress(object sender, KeyPressEventArgs e)
        {
            (_processor as ReportEffectivenessProcessor).SetFormula();
        }

        private void DgvOtclkInfrorm_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportInfrormationResponseProcessor).SetFormula();
        }

        private void DgvCadre_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportCadreProcessor).SetFormula();
        }

        private void DgvReqVCR_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportReqVCRProcessor).SetFormula();
        }

        private void dgvOpedQ_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportOpedQProcessor).SetCalculateValue();

        }

        private void gVac_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportVaccinationProccesor).SetFormulaMonth();
        }

        private void DgwReportIizl_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dgvFssM_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {

        }

        private void dgvFssM_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as FSSMonitoringProcessor).SetFormula();

        }

        private void dgvEffectiveness_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //(_processor as ReportEffectivenessProcessor).SetFormula();

        }

        private void сводToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.ConsolidateVSS);
        }

        private void dgvMonitoringVCR_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            (_processor as MonitoringVCRProcessor).SetFormula();
        }

        private void dgvMonitoringVCR_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as MonitoringVCRProcessor).SetFormula();

        }

        private void dgvMonitoringVCR_KeyPress(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as MonitoringVCRProcessor).SetFormula();

        }

        private void dgvQuantity_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            (_processor as ReportQuantityProcessor).SetFormula();
        }

        private void dgvQuantity_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportQuantityProcessor).SetFormula();

        }

        private void dgvTarAllow_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportTargetedAllowancesProcessor).SetFormula();

        }

        private void dgvPVPLoad_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportPVPLoadProcessor).SetFormula();

        }

        private void dgvDoff_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportDoffProcessor).SetFormula();

        }


        private void dgvTarAllow_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            (_processor as ReportTargetedAllowancesProcessor).SetFormula();

        }


        private void dgvPVPLoad_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            (_processor as ReportPVPLoadProcessor).SetFormula();

        }

        private void dgvDoff_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            (_processor as ReportDoffProcessor).SetFormula();

        }


        private void DgvReqVCR_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            (_processor as ReportReqVCRProcessor).SetFormula();
        }

        private void DgvReqVCR_KeyPress(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportReqVCRProcessor).SetFormula();

        }


        private void сводToolStripMenuItemVCR_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.ConsolidateVCR);
        }

        private void сводПоФилиаламToolStripMenuItemVCRFilial_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.ConsolidateVCRFilial);
        }

        private void сводПоВсемФилиаламToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.ConsQuantityFilial);
        }

        private void сводToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.ConsolidateOpedQ);
        }

        private void сводToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.ConsolidateOpedUnplanned);
        }

        private void dgvProposal_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            (_processor as ReportProposalProcessor).CalculateCells();

        }

        private void dgvOpedFinance_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportOpedFinanceProcessor).CalculateCells();
        }

        private void dgvOpedFinance3_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportOpedFinance3Processor).CalculateCells();
        }

        private void цПНП2квартальныйToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.ConsolidateCPNP2Q);
        }

        private void планРезультативностиЭкспертнойДеятельностиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var form = new OpedFinanceDictionaryForm(_client);
            form.Show();
            form.FormClosed += (s, ee) =>
            {
                if (form != null)
                {
                    if (!form.IsDisposed)
                        form.Dispose();

                    form = null;
                }
            };
        }

        private void планЧисленностиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var form = new QuantityPlanDictionaryForm(_client);
            form.Show();
            form.FormClosed += (s, ee) =>
            {
                if (form != null)
                {
                    if (!form.IsDisposed)
                        form.Dispose();

                    form = null;
                }
            };
        }

        private void свод1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.ConsOpedFinance1);
        }

        private void cbIizl2022_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangeIndexComboBox(dgvIizl2022, cbIizl2022, tbIizl2022);
        }

        private void dgvIizl2022_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            (_processor as ReportIizlProcessor2022).SetCalculateCellsValue();
        }

        private void свод2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.ConsOpedFinance2);
        }

        private void свод3ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.ConsOpedFinance3);
        }

        private void DgvQuery_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            _dynamicReportProcessor.CalculateCells();
        }

        private void сводToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            OpenConsolidateReportForm(ConsolidateReport.ConsPropsal);
        }

        private void DgwReportIizl_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}