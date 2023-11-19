using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Serialization;
using KmsReportClient.Excel.Creator;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Model;
using KmsReportClient.Model.XML;
using KmsReportClient.Report.Common;
using KmsReportClient.Support;

namespace KmsReportClient.Report.Basic
{
    public class DynamicReportProcessor
    {
        private readonly EndpointSoap client;
        public List<DynamicDataDto> data = new List<DynamicDataDto>();
        public DynamicReport Report;
        private ComboBox _cb;
        private DataGridView Dgv;
        public string _pageName;
        public int _pageIndex;
        public int oldPageIndex;
        private ReportDynamicDto ReportDynamic;
        private ReportDynamicFlowDto ReportDynamicFlow;

        //Для проверки 2022
        private CheckFFOMS2022Common _checkFFOMS2022Common { get; set; }
        private CheckFFOMS2022Common CheckFFOMS2022Common {
            get {
                if (_checkFFOMS2022Common == null && Report != null && Report.Id == 33)
                    return new CheckFFOMS2022Common(client);
                else return _checkFFOMS2022Common;
            }

        }

        public DynamicReportProcessor(EndpointSoap client, DataGridView dgv, ComboBox cb)
        {
            this.client = client;
            this.Dgv = dgv;
            this._cb = cb;
        }

        public DynamicReportProcessor(EndpointSoap client)
        {
            this.client = client;

        }

        public void TuneProverkaTfomsTables(string year, string idRegion)
        {
            if (_cb.Text.ToLower().Contains("леталь"))
            {

                var commonLetalData = CheckFFOMS2022Common.GetFFOMS2022CommonData(year, idRegion);

                #region Общее количество экспертиз
                var parseNum = GlobalUtils.TryParseDecimal(commonLetalData.CountLetalAll);
                string currentPosition = $"P{_cb.SelectedIndex.ToString()}C0R0";
                var currentNum = data.FirstOrDefault(x => x.Position == currentPosition);

                if (parseNum != 0 && currentNum == null)
                {
                    Dgv.Rows[0].Cells[0].Value = Convert.ToInt32(parseNum);

                }

                #endregion

                #region Количество ЭКМП
                parseNum = GlobalUtils.TryParseDecimal(commonLetalData.CountEkmp);
                currentPosition = $"P{_cb.SelectedIndex.ToString()}C1R0";
                currentNum = data.FirstOrDefault(x => x.Position == currentPosition);

                if (parseNum != 0 && currentNum == null)
                {
                    Dgv.Rows[0].Cells[1].Value = Convert.ToInt32(parseNum);

                }

                #endregion

                #region Количество нарушений

                parseNum = GlobalUtils.TryParseDecimal(commonLetalData.CountNarush);
                currentPosition = $"P{_cb.SelectedIndex.ToString()}C3R0";
                currentNum = data.FirstOrDefault(x => x.Position == currentPosition);
                if (parseNum != 0 && currentNum == null)
                {
                    Dgv.Rows[0].Cells[3].Value = Convert.ToInt32(parseNum);

                }

                #endregion

                #region Количество непроведённых по обоснованной причине

                parseNum = GlobalUtils.TryParseDecimal(commonLetalData.CountNeProvedenaOb);
                currentPosition = $"P{_cb.SelectedIndex.ToString()}C5R0";
                currentNum = data.FirstOrDefault(x => x.Position == currentPosition);
                if (parseNum != 0 && currentNum == null)
                {
                    Dgv.Rows[0].Cells[5].Value = Convert.ToInt32(parseNum);

                }

                #endregion

                Dgv.Columns[2].ReadOnly = Dgv.Columns[4].ReadOnly = Dgv.Columns[7].ReadOnly = true;

            }
            else
            {

                if (_cb.Text.ToLower().Contains("экмп"))
                {
                    var commonLetalData = CheckFFOMS2022Common.GetFFOMS2022CommonData(year, idRegion);

                    #region Общее количество экспертиз
                    var parseNum = GlobalUtils.TryParseDecimal(commonLetalData.CountLetalAll);
                    string currentPosition = $"P{_cb.SelectedIndex.ToString()}C2R1";
                    var currentNum = data.FirstOrDefault(x => x.Position == currentPosition);

                    if (parseNum != 0 && currentNum == null)
                    {
                        Dgv.Rows[1].Cells[2].Value = Convert.ToInt32(parseNum);

                    }

                    #endregion

                    #region Количество ЭКМП
                    parseNum = GlobalUtils.TryParseDecimal(commonLetalData.CountEkmp);
                    currentPosition = $"P{_cb.SelectedIndex.ToString()}C3R1";
                    currentNum = data.FirstOrDefault(x => x.Position == currentPosition);

                    if (parseNum != 0 && currentNum == null)
                    {
                        Dgv.Rows[1].Cells[3].Value = Convert.ToInt32(parseNum);

                    }

                    #endregion



                }

                Dgv.Columns[4].ReadOnly = true;
            }

            CalculateCells();
        }

        public void ToExcel(string fileName)
        {
            DynamicReportExcelCreator excel = new DynamicReportExcelCreator(fileName, Report);
            excel.CreateReport();
        }

        public void ToExcel(string fileName, List<DynamicDataDto> data)
        {

            DynamicReportExcelCreator excel = new DynamicReportExcelCreator(fileName, Report, data);
            excel.CreateReport();
        }


        public void SetReadOnlyDgv(DataGridView dgv, int idFlow)
        {
            dgv.ReadOnly = false;
            if (!Report.IsUserRow)
            {
                dgv.AllowUserToAddRows = false;
            }
            else
            {
                dgv.AllowUserToAddRows = true;

            }

            if (CurrentUser.IsMain)
            {
                dgv.ReadOnly = true;
            }

            var flow = GetReportDynamicFlow(idFlow);
            if (flow != null)
            {
                if (flow.Status.ToString() == "Submit" || flow.Status.ToString() == "Done")
                    dgv.ReadOnly = true;
            }


        }

        public ReportDynamicFlowDto GetReportDynamicFlow(int idFlow)
        {
            var request = new GetReportDynamicFlowByIdRequest
            {
                Body = new GetReportDynamicFlowByIdRequestBody
                {
                    flowId = idFlow
                }
            };
            return client.GetReportDynamicFlowById(request).Body.GetReportDynamicFlowByIdResult;
        }



        public ReportDynamicDto GetReportDynamic(int idReport)
        {
            var request = new GetReportDynamicByIdRequest
            {
                Body = new GetReportDynamicByIdRequestBody
                {
                    reportId = idReport
                }
            };
            return client.GetReportDynamicById(request).Body.GetReportDynamicByIdResult;
        }

        public void SetReportDynamic(int reportId)
        {
            ReportDynamic = GetReportDynamic(reportId);
        }

        public void SetReportDynamicFlow(int flowId)
        {
            ReportDynamicFlow = GetReportDynamicFlow(flowId);

        }



        public void SetData(DataGridView dgv, int pageIndex)
        {

            if (data.Count != 0)
            {
                data.RemoveAll(x => PositionSupport.GetPage(x.Position) == pageIndex);
            }

            var oldTheme = new List<DynamicDataDto>();
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                int startColumn = 2;
                if (dgv.Columns[0].Name != "Наименование показателя")
                    startColumn = 0;

                for (int j = startColumn; j < dgv.Columns.Count; j++)
                {
                    var cell = data.Where(x => x.Position == GetPosition(pageIndex, j, i)).FirstOrDefault();
                    if (cell == null)
                    {

                        if (dgv.Rows[i].Cells[j].Value != null)
                        {
                            oldTheme.Add(new DynamicDataDto
                            {
                                Position = GetPosition(pageIndex, j, i),
                                Value = dgv.Rows[i].Cells[j].Value == null ? "" : Convert.ToString(dgv.Rows[i].Cells[j].Value)
                            });
                        }

                    }
                    else
                    {
                        cell.Value = dgv.Rows[i].Cells[j].Value == null ? "" : Convert.ToString(dgv.Rows[i].Cells[j].Value);
                    }

                }
            }

            data.AddRange(oldTheme);

        }

        public string GetReportInfo(int idFlow)
        {
            var flow = GetReportDynamicFlow(idFlow);

            string info = $"{ReportDynamic.Name.Trim()}; Дата: {ReportDynamic.Date.ToShortDateString()}; " + Environment.NewLine;

            if (!string.IsNullOrEmpty(ReportDynamic.Description.Trim()))
            {
                info += $"Описание: {ReportDynamic.Description.Trim()}" + Environment.NewLine;
            }
            info += $"Создал: {CurrentUser.Users.First(x => Convert.ToInt32(x.Key) == ReportDynamic.UserCreated).Value}" + Environment.NewLine;
            if (flow != null)
            {
                var statusText = GlobalConst.FilterList.First(x => x.Key == flow.Status.ToString()).Value;
                info += $"Статус: {statusText}" + Environment.NewLine;
            }


            return info;
        }

        public void FillThemeData(DataGridView dgv)
        {
            var themeData = data.Where(x => PositionSupport.GetPage(x.Position) == _pageIndex).ToList();

            if (themeData.Count == 0)
            {
                return;
            }
            foreach (var cell in themeData)
            {
                int row = GetRow(cell.Position);
                int col = GetColumn(cell.Position);
                dgv.Rows[row].Cells[col].Value = cell.Value;
            }

        }

        public int GetColumn(string position)
        {
            char[] delimiterChars = { 'P', 'R', 'C' };
            string[] stringSplit = position.Split(delimiterChars);
            return Convert.ToInt32(stringSplit[2]);
        }

        public int GetRow(string position)
        {
            char[] delimiterChars = { 'P', 'R', 'C' };
            string[] stringSplit = position.Split(delimiterChars);
            return Convert.ToInt32(stringSplit[3]);
        }

        public int SaveReportFiliialData()
        {
            if (data.Count == 0)
            {
                MessageBox.Show("Ошибка! Внесите данные!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 0;
            }

            var request = new SaveDynamicFlowDataRequest
            {
                Body = new SaveDynamicFlowDataRequestBody
                {
                    data = data.ToArray(),
                    fillialCode = CurrentUser.FilialCode,
                    IdReportDynamic = Report.Id,
                    idUser = CurrentUser.IdUser
                }
            };
            MessageBox.Show("Данные успешно Сохранены!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return client.SaveDynamicFlowData(request).Body.SaveDynamicFlowDataResult;



        }

        public List<DynamicDataDto> GetRegionData(int idFlow)
        {
            var request = new GetReportRegionDataRequest
            {
                Body = new GetReportRegionDataRequestBody
                {
                    idFlow = idFlow
                }
            };
            return client.GetReportRegionData(request).Body.
                 GetReportRegionDataResult.
                 ToList();
        }

        public string GetPosition(int page, int column, int row) => String.Format($"P{page}C{column}R{row}");


        public List<ReportDynamicDto> GetDynamicReports() =>
            client.GetDynamicReports(new GetDynamicReportsRequest()).Body.GetDynamicReportsResult.ToList();

        public GetDynamicReportResponse GetXmlReport(int reportId)
        {
            var request = new GetDynamicReportXmlRequest(new GetDynamicReportXmlRequestBody
            {
                reportId = reportId
            });
            return client.GetDynamicReportXml(request).Body.GetDynamicReportXmlResult;
        }


        public void SetReport(GetDynamicReportResponse reportResponse)
        {

            var deserializer = new XmlSerializer(typeof(TemplateDynamicReport));
            TemplateDynamicReport Xml;
            using (TextReader reader = new StringReader(reportResponse.Xml))
            {
                Xml = (TemplateDynamicReport)deserializer.Deserialize(reader);
            }

            if (Xml == null)
            {
                MessageBox.Show("Ошибка загрузки отчёта");

            }

            data.Clear();
            Report = new DynamicReport(Xml, reportResponse.Id);

        }



        public void SetDgv(DataGridView dgv, string page)
        {
            int ColCounter = 0;
            dgv.Rows.Clear();
            dgv.Columns.Clear();
            int dgvColumnCount = dgv.Columns.Count;
            string columnName;

            


            foreach (var theme in Report.Page.Where(x => x.Key.Name == page))
            {
                ColCounter = theme.Value.Columns.Count();
                //Создаём столбцы
                foreach (var column in theme.Value.Columns)
                {
                    columnName = CreateColumnName("", column.Name);

                    if (column.IsGroup)
                    {
                        foreach (var childColumn in column.Columns)
                        {
                            columnName = CreateColumnName(column.Name, childColumn.Name);
                            dgv.Columns.Add("Column" + dgvColumnCount, columnName);
                        }
                    }
                    else
                    {
                        dgv.Columns.Add("Column" + dgvColumnCount, columnName);
                    }
                }

                int RowCounter = (int) Math.Ceiling((decimal) data.Count() / (decimal) ColCounter);

                //Добавляем строки
                if (theme.Value.Rows.Any())
                {
                    for (int i = 0; i < RowCounter; i++)
                        dgv.Rows.Add();
                }
                else
                {
                    dgv.Rows.Add();
                }
            }
        }



        public string GetDescriptionPage(string page)
        {

            return Report.Page.Where(p => p.Key.Name == page).FirstOrDefault().Key.Description ?? "";
        }

        public string CreateColumnName(string group, string column) => String.Format($"{group};{column}");

        public void SetComboBox(ComboBox cmb)
        {
            var PageData = Report.Page.Select(x => new { x.Key.Index, x.Key.Name, x.Key.Description });
            cmb.DisplayMember = "Name";
            cmb.ValueMember = "Index";
            cmb.DataSource = PageData.ToList();

        }

        internal void CalculateCells()
        {
            if (Report.Id == 33 && _cb.Text.ToLower().Contains("летал"))
            {

                var row = Dgv.Rows[0];
                try
                {

                    row.Cells[2].Value = Math.Round(GlobalUtils.TryParseDecimal(row.Cells[1].Value) / GlobalUtils.TryParseDecimal(row.Cells[0].Value) * 100, 2);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                }


                try
                {

                    row.Cells[4].Value = Math.Round(GlobalUtils.TryParseDecimal(row.Cells[3].Value) / GlobalUtils.TryParseDecimal(row.Cells[1].Value) * 100, 2);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);

                }
                try
                {

                    row.Cells[7].Value = Math.Round(GlobalUtils.TryParseDecimal(row.Cells[6].Value) / GlobalUtils.TryParseDecimal(row.Cells[0].Value) * 100, 2);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);

                }

            }
            else if (Report.Id == 33)
            {
                foreach (DataGridViewRow row in Dgv.Rows)
                {
                    try
                    {
                        if (GlobalUtils.TryParseDecimal(row.Cells[2].Value) != 0)
                        {
                            row.Cells[4].Value = Math.Round(GlobalUtils.TryParseDecimal(row.Cells[3].Value) / GlobalUtils.TryParseDecimal(row.Cells[2].Value) * 100, 2);

                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);

                    }
                }

            }
        }
    }
}
