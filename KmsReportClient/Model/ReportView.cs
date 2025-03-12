using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Report.Basic;
using Microsoft.Office.Interop.Word;
using NLog;

namespace KmsReportClient.Model
{
    public class ReportView
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        private const string Monthly = "monthly";
        private const string Quarterly = "quarterly";
        private const string FFOMS = "ffoms";

        private readonly TreeView tree;
        private readonly List<KmsReportDictionary> regions;

        private readonly EndpointSoap client;
        private readonly DynamicReportProcessor dynamicReportProcessor;

        private readonly Dictionary<string, string> rootDict =
            new Dictionary<string, string> { { Monthly, "Ежемесячные отчеты" }, { Quarterly, "Квартальные отчеты" }, { FFOMS, "К проверке ФФОМС" } };

        private readonly List<ReportHistory> reportsHistory;

        public ReportView(
            TreeView tree,
            List<KmsReportDictionary> regions,
            List<KmsReportDictionary> reportsDictionary,

            EndpointSoap client)
        {
            this.regions = regions;
            this.client = client;
            this.tree = tree;

            reportsHistory = CreateHistoryList(reportsDictionary);
            dynamicReportProcessor = new DynamicReportProcessor(client);


        }

        public bool CreateTreeView(int year)
        {
            try
            {
                var flows = CollectFlow(year);
                int yymmEnd = Convert.ToInt32(year.ToString().Substring(2) + "12");
                tree.Nodes.Clear();

                foreach (var root in rootDict)
                {
                    var rootNode = new TreeNode { Text = root.Value };
                    tree.Nodes.Add(rootNode);

                    if (root.Key == FFOMS)
                    {
                        // Обработка FFOMS: отчеты без даты
                        var ffomsReports = reportsHistory.Where(x => x.Type == FFOMS);
                        foreach (var report in ffomsReports)
                        {
                            var reportNode = new TreeNode { Text = report.HistoryName };
                            rootNode.Nodes.Add(reportNode);
                            if (CurrentUser.IsMain)
                            {
                                // Для главного администратора: добавляем все регионы
                                foreach (var region in regions)
                                {
                                    var regionNode = new TreeNode { Text = region.Value };
                                    var date = DateTime.ParseExact("2503", "yyMM", CultureInfo.InvariantCulture);
                                    var color = FindReportColor(date, report.HistoryCode, region.Key, flows);
                                    reportNode.Nodes.Add(regionNode);
                                    if (color != null) regionNode.BackColor = color.Value;
                                    
                                }
                            }
                            else
                            {
                                // Для обычного пользователя: только текущий филиал
                                var date = DateTime.ParseExact(report.YymmEnd, "yyMM", CultureInfo.InvariantCulture);
                                var color = FindReportColor(date, report.HistoryCode, CurrentUser.FilialCode, flows);
                                if (color != null) reportNode.BackColor = color.Value;
                            }
                        }
                    }
                    else
                    {
                        var dates = CreateDateList(root.Key, year);
                        foreach (var history in reportsHistory.Where(x => x.Type == root.Key))
                        {
                            //Берём только те отчёты, которые нужно показать
                            if (Convert.ToInt32(history.YymmEnd) < yymmEnd)
                                continue;

                            //Тот кто будет читать простите))
                            if (history.HistoryCode == "iizl2022" && Convert.ToInt32(yymmEnd) <= 2112)
                                continue;

                            var nodeReport = new TreeNode { Text = history.HistoryName };
                            rootNode.Nodes.Add(nodeReport);

                            foreach (var date in dates)
                            {
                                var nodePeriod = new TreeNode { Text = date.ToString("MMMM yyyy") };
                                nodeReport.Nodes.Add(nodePeriod);

                                if (!CurrentUser.IsMain)
                                {
                                    var color = FindReportColor(date, history.HistoryCode, CurrentUser.FilialCode, flows);
                                    if (color != null)
                                    {
                                        nodePeriod.BackColor = color.Value;
                                    }
                                }
                                else
                                {
                                    foreach (var region in regions)
                                    {
                                        var nodeRegion = new TreeNode { Text = region.Value };
                                        nodePeriod.Nodes.Add(nodeRegion);
                                        var color = FindReportColor(date, history.HistoryCode, region.Key, flows);
                                        if (color != null)
                                        {
                                            nodeRegion.BackColor = color.Value;
                                        }
                                    }

                                }
                            }
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error creating report tree");
                throw;
            }
        }

        public bool CreateTreeViewQuery(int year)
        {
            try
            {
                var ReportList = dynamicReportProcessor.GetDynamicReports();
                var Months = CreateDateList(Monthly, year);
                var flows = CollectFlowDynamic(year);

                foreach (var month in Months)
                {
                    string monthString = month.ToString("MMMM yyyy");
                    TreeNode monthNode = tree.Nodes.Add(monthString);
                    var ReportsMonth = ReportList.Where(x => x.Date.Year == month.Year && x.Date.Month == month.Month).ToList();

                    foreach (var reportMonth in ReportsMonth)
                    {
                        var reportTag = new ReportNodeTag { IdReport = reportMonth.Id };
                        var reportNode = new TreeNode { Text = reportMonth.Name,Tag = reportTag };
                        monthNode.Nodes.Add(reportNode);

                        if (CurrentUser.IsMain)
                        {
                            foreach (var region in regions)
                            {
                                var regionNode = reportNode.Nodes.Add(region.Value);
                                var flow = flows.FirstOrDefault(x => x.IdReport == reportMonth.Id && region.Key == x.IdRegion);
                                var tag = new ReportNodeTag { IdReport = reportMonth.Id };
                                regionNode.Tag = tag;
                                if (flow != null)
                                {
                                    regionNode.BackColor = GetColorForNode(flow.Status).Value;
                                    tag.idFlow = flow.IdFlow;


                                }

                            }
                        }
                        else
                        {
                            var flow = flows.FirstOrDefault(x => x.IdReport == reportMonth.Id && CurrentUser.FilialCode == x.IdRegion);
                            if (flow != null)
                            {

                                reportNode.BackColor = GetColorForNode(flow.Status).Value;
                                reportTag.idFlow = flow.IdFlow;
                                     
                            }




                        }

                    }

                }
                return true;
                // return !CurrentUser.IsMain && flows.Any(x => x.Status == ReportStatus.Refuse);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error creating report tree");
                throw;
            }
        }

        private ReportFlowDto[] CollectFlow(int year)
        {
            string filial = CurrentUser.IsMain ? "" : CurrentUser.FilialCode;
            var yymmEnd = DateTime.Today.ToString("yyMM");
            if (year != DateTime.Today.Year)
            {
                yymmEnd = new DateTime(year, 12, 15).ToString("yyMM");
            }

            var yymmStart = yymmEnd.Substring(0, 2) + "01";
            var request = new GetFlowsRequest() { Body = new GetFlowsRequestBody(filial, yymmStart, yymmEnd) };
            return client.GetFlows(request).Body.GetFlowsResult;
        }

        private ReportDynamicFlowDto[] CollectFlowDynamic(int year)
        {
            var request = new GetReportDynamicFlowsRequest { Body = new GetReportDynamicFlowsRequestBody { year = year } };
            return client.GetReportDynamicFlows(request).Body.GetReportDynamicFlowsResult;
        }

        private List<DateTime> CreateDateList(string type, int year)
        {

            int currentMonth = DateTime.Today.Month;

            if (year != DateTime.Today.Year)
            {
                currentMonth = 12;

            }

            var dates = new List<DateTime>();

            for (int i = currentMonth - 1; i >= 0; i--)
            {
                //var date = DateTime.Today.AddMonths(-i);
                var date = new DateTime(year, currentMonth - i, 15);

                if (type == Quarterly)
                {
                    int ost = date.Month % 3;
                    if (ost != 0)
                    {
                        continue;
                    }
                }

                dates.Add(date);
            }

            return dates;
        }

        private List<ReportHistory> CreateHistoryList(List<KmsReportDictionary> reportsDictionary) =>
            reportsDictionary.Select(report =>
                    new ReportHistory { HistoryName = report.Value, HistoryCode = report.Key, Type = report.ForeignKey, YymmEnd = report.AdditionalField })
                .ToList();

        private Color? FindReportColor(DateTime date, string historyCode, string idRegion, ReportFlowDto[] flows)
        {
            var status = flows.SingleOrDefault(x =>
                    x.IdRegion == idRegion && x.Yymm == date.ToString("yyMM") && x.IdReport == historyCode)
                ?.Status;

            return GetColorForNode(status);
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
    }

    public class ReportHistory
    {
        public string HistoryName { get; set; }
        public string HistoryCode { get; set; }
        public string Type { get; set; }

        public string YymmEnd { get; set; }
    }
}