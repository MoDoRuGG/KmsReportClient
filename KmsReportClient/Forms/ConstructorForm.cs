using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Model;
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.XML;
using KmsReportClient.Report.Basic;
using KmsReportClient.Support;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace KmsReportClient.Forms
{
    public partial class ConstuctorForm : Form
    {
        private const string StrConst = "Строки";
        private DynamicReport _report;
        private string _currentPage;
        private int _currentPageIndex;
        private readonly EndpointSoapClient _client;
        private DynamicReportProcessor processor;
        int reportId;
        private bool _isEdit = false;

        public ConstuctorForm(EndpointSoapClient client)
        {
            InitializeComponent();
            _client = client;
            _report = new DynamicReport();
            CbxPage.DataSource = _report.Page.Keys.ToList();
            CbxPage.DisplayMember = "Name";
            CbxPage.ValueMember = "Name";
            _currentPage = CbxPage.Items[0].ToString();
            _report.NameReport = "Новый отчёт";
            _report.DateReport = DateTime.Now;


            CbxShow.DataSource = new List<string> { "Столбцы", "Строки" };
            CbxShow.SelectedIndex = 0;

            LbEmail.DataSource = _client.GetEmails();
            LbEmail.DisplayMember = "ForeignKey";
            LbEmail.ValueMember = "Key";

            TbxTabDesc.Text = "1 вкладка";


            SetDgProperties();
            CreateTreeListView();
            this.treeListView1.AutoResizeColumns();

        }


        public ConstuctorForm(EndpointSoapClient client, int idReport)
        {
            InitializeComponent();
            _client = client;
            this.reportId = idReport;
            processor = new DynamicReportProcessor(client);
            var xml = processor.GetXmlReport(idReport);
            processor.SetReport(xml);

            _report = processor.Report;
            TbxName.Text = _report.NameReport;
            TbxDescReport.Text = _report.DescriptionReport;
            //   DtmDate.Value = _report.DateReport;

            CbxUserRow.Checked = _report.IsUserRow;
            LbEmail.DataSource = _client.GetEmails();
            LbEmail.DisplayMember = "ForeignKey";
            LbEmail.ValueMember = "Key";

            for (int i = 0; i < LbEmail.Items.Count; i++)
            {
                var itemEmail = _report.Executors.Where(x => x == (LbEmail.Items[i] as KmsReportDictionary).Value).FirstOrDefault();
                if (itemEmail != null)
                    LbEmail.SetItemChecked(i,true);
            }



            CbxPage.DataSource = _report.Page.Keys.ToList();
            CbxPage.DisplayMember = "Name";
            CbxPage.ValueMember = "Name";
            _currentPage = CbxPage.Items[0].ToString();
            _report.NameReport = "Новый отчёт";
            _report.DateReport = DateTime.Now;

            CbxShow.DataSource = new List<string> { "Столбцы", "Строки" };
            CbxShow.SelectedIndex = 0;
            SetDgProperties();
            CreateTreeListView();
            CbxChangeRowsColumns();
            this.treeListView1.AutoResizeColumns();
        }

        private void dg_apper_property_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            группаToolStripMenuItem.Visible = true;
            if (CbxShow.SelectedItem.Equals("Строки"))
            {
                using var createColumnForm = new CreateElementForm(_report, _currentPageIndex, TableElement.Row);
                createColumnForm.ShowDialog();
            }
            else
            {
                var SelectedItem = treeListView1.SelectedObject as Group ?? new Group();
                var Parent = treeListView1.GetParent(SelectedItem);
                if (Group.IsSubItem(Parent) || SelectedItem.IsGroup)
                {
                    группаToolStripMenuItem.Visible = false;
                }

                btnAddContextMenuStrip.Show(sender as Button, new Point(0, 0));


            }
            CbxChangeRowsColumns();
        }

        public void CbxChangeRowsColumns()
        {
            if (CbxShow.SelectedItem is null)
                return;

            treeListView1.Roots = null;
            string CbxShowValue = CbxShow.SelectedItem.ToString();
            var ReportData = _report.Page.ElementAt(_currentPageIndex);
            ReportData.Value.ReIndexItems();
            if (CbxShowValue.Equals("Строки"))
            {

                treeListView1.Roots = ReportData.Value.Rows;
            }
            else
            {
                treeListView1.Roots = ReportData.Value.Columns;
            }

        }

        public void CreateTreeListView()
        {
            treeListView1.CanExpandGetter = model => ((Group)model).
                                                          Columns.Count > 0;
            treeListView1.ChildrenGetter = delegate (object model)
            {
                return ((Group)model).Columns;
            };

            var IndexCol = new BrightIdeasSoftware.OLVColumn("Номер", "Index");
            IndexCol.AspectGetter = delegate (object x) { return ((Group)x).Index; };

            var NameCol = new BrightIdeasSoftware.OLVColumn("Наименование", "Name");
            NameCol.AspectGetter = delegate (object x) { return ((Group)x).Name; };

            this.treeListView1.Columns.Add(IndexCol);
            this.treeListView1.Columns.Add(NameCol);


        }

        private void CbxShow_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CbxShow.SelectedIndex == 0)
            {
                BtnAdd.ToolTipText = "Создать столбец";
                Группа.Enabled = true;
            }
            else
            {
                BtnAdd.ToolTipText = "Создать строку";
                Группа.Enabled = false;
            }

            CbxChangeRowsColumns();
        }

        private void BtnCreatePage_Click(object sender, EventArgs e)
        {
            using var form = new CreateElementForm(CbxPage, _report, _currentPageIndex, TableElement.Page);
            form.ShowDialog();
            CbxPage.SelectedIndex = CbxPage.Items.Count - 1;
        }

        private void CbxPage_SelectedIndexChanged(object sender, EventArgs e)
        {
            _currentPage = CbxPage.SelectedItem.ToString();
            _currentPageIndex = CbxPage.SelectedIndex;

            //Console.WriteLine(_currentPageIndex);
            //Console.WriteLine(_currentPage);

            var page = _report.Page.ElementAt(_currentPageIndex).Key;
            TbxTabDesc.Text = page.Description;

            TbxNameElement.Text = page.Name;
            TbxDescElement.Text = page.Description;
            CbxChangeRowsColumns();

        }

        private void SetDgProperties()
        {

        }

        private void BtnSave_Click(object sender, EventArgs e)
        {

        }

        public bool ValidFiled()
        {
            if (TbxName.Text.Trim() == String.Empty)
            {
                MessageBox.Show("Введите наименование отчёта!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (TbxName.Text.Trim().Length <= 4)
            {
                MessageBox.Show("Наименование отчёта слишком короткое", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }


            foreach (var page in _report.Page.Values)
            {
                if (!page.Columns.Any())
                {
                    MessageBox.Show("Остались незаполненные вкладки", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;

                }
            }

            return true;
        }

        public bool ValidElementField()
        {
            if (TbxNameElement.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Имя элемента не может быть пустым!", "Редактирование", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }


        private void Save()
        {
            if (!ValidFiled())
            {
                return;
            }

            _report.NameReport = TbxName.Text.Trim();
            _report.DateReport = DtmDate.Value;
            _report.DescriptionReport = TbxDescReport.Text.Trim();
            _report.IsUserRow = CbxUserRow.Checked;


            _report.Executors.Clear();
            foreach (var item in LbEmail.CheckedItems)
            {
                _report.Executors.Add((item as KmsReportDictionary).Value);
            }

            var xml = new TemplateDynamicReport();
            xml.Name = _report.NameReport;
            xml.ReportDate = _report.DateReport;
            xml.ReportDescription = _report.DescriptionReport;
            xml.tables = new List<TemplateTableDynamic>();          
            xml.Executors = _report.Executors;
            xml.IsUserRow = _report.IsUserRow;

            foreach (var page in _report.Page)
            {
                var templateTableDynamic = new TemplateTableDynamic();
                templateTableDynamic.Name = page.Key.Name;
                templateTableDynamic.TableDescription = page.Key.Description;
                templateTableDynamic.Columns = new List<TemplateColumnDynamic>();
                templateTableDynamic.Rows = new List<TemplateRowDynamic>();

                foreach (var col in page.Value.Columns)
                {
                    var column = new TemplateColumnDynamic();
                    column.ColumnDescription = col.Description;
                    column.NameColumn = col.Name;
                    column.IndexColumn = col.Index;
                    if (col.IsGroup)
                    {
                        column.ChildColumn = col.Columns.Select(c => new TemplateColumnDynamic
                        {
                            NameColumn = c.Name,
                            ColumnDescription = c.Description,
                            IndexColumn = c.Index.ToString()

                        }).ToList();

                    }
                    templateTableDynamic.Columns.Add(column);
                }

                foreach (var row in page.Value.Rows)
                {
                    var templateRow = new TemplateRowDynamic();
                    templateRow.NameRow = row.Name;
                    templateRow.RowDescription = row.Description;
                    templateRow.IndexRow = row.Index;
                    templateTableDynamic.Rows.Add(templateRow);
                }

                xml.tables.Add(templateTableDynamic);

            }

            var report = CreateReport();
            using (var stringwriter = new System.IO.StringWriter())
            {
                string reportName = report.FileName;
                var serializer = new XmlSerializer(typeof(TemplateDynamicReport));
                serializer.Serialize(stringwriter, xml);
                _client.UploadXmlDynamicFile(Encoding.UTF8.GetBytes(stringwriter.ToString().ToCharArray()), reportName);
            }

            MessageBox.Show("Отчётная форма успешно сохранена.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Close();


        }

        private Report_Dynamic CreateReport()
        {
            try
            {
                var report = new ReportDynamicDto
                {
                    Id = reportId,
                    Date = _report.DateReport,
                    Description = _report.DescriptionReport,
                    FileName = String.Format($"{ _report.NameReport }_{ _report.DateReport.ToShortDateString().Replace(".", "") }.xml"),
                    Name = _report.NameReport,
                    UserCreated = CurrentUser.IdUser,
                    IsUserRow = _report.IsUserRow
                };
                return _client.CreateDynamicReport(report);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }

        }

        private void BtnAddGroup_Click(object sender, EventArgs e)
        {

        }

        private void столбецToolStripMenuItem_Click(object sender, EventArgs e)
        {

            var SelectedItem = treeListView1.SelectedObject as Group;
            using var createColumnForm = new CreateElementForm(SelectedItem, _report, _currentPageIndex, TableElement.Column);
            createColumnForm.ShowDialog();
            CbxChangeRowsColumns();
        }

        private void группаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using var form = new CreateElementForm(_report, _currentPageIndex, TableElement.Group);
            form.ShowDialog();
            CbxPage.SelectedIndex = CbxPage.Items.Count - 1;
            CbxChangeRowsColumns();
        }

        private void BtnAdd_MouseDown(object sender, MouseEventArgs e)
        {

        }
        private void снятьВыделениеToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }


        private void ConstructorForm_Load(object sender, EventArgs e)
        {
            TbxDescElement.Enabled = false;
            TbxNameElement.Enabled = false;
        }

        private void BtnUp_Click(object sender, EventArgs e)
        {

        }



        private void btnAddContextMenuStrip_Opening(object sender, CancelEventArgs e)
        {

        }

        private void treeListView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void treeListView1_SelectionChanged(object sender, EventArgs e)
        {
            var group = treeListView1.SelectedObject as Group;
            if (group == null)
                return;
            TbxNameElement.Text = group.Name;
            TbxDescElement.Text = group.Description;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (CbxShow.SelectedItem.Equals("Строки"))
            {
                using var createColumnForm = new CreateElementForm(_report, _currentPageIndex, TableElement.Row);
                createColumnForm.ShowDialog();
            }
            else
            {
                var SelectedItem = treeListView1.SelectedObject;
                if (SelectedItem != null && treeListView1.GetParent(SelectedItem) != null)
                    SelectedItem = treeListView1.GetParent(SelectedItem);
                using var createColumnForm = new CreateElementForm(SelectedItem, _report, _currentPageIndex, TableElement.Column);
                createColumnForm.ShowDialog();
            }

            CbxChangeRowsColumns();

            //this.treeListView1.AutoResizeColumns();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            using var form = new CreateElementForm(_report, _currentPageIndex, TableElement.Group);
            form.ShowDialog();
            CbxPage.SelectedIndex = CbxPage.Items.Count - 1;
            CbxChangeRowsColumns();
            //   this.treeListView1.AutoResizeColumns();
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void BtnSaveReport_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void BtnEditElement_Click(object sender, EventArgs e)
        {

        }

        private void TbxNameElement_TextChanged(object sender, EventArgs e)
        {

        }

        private void TbxDescElement_TextChanged(object sender, EventArgs e)
        {

        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {

            var group = treeListView1.SelectedObject as Group;
            if (group == null)
                return;
            var page = _report.Page.First(x => x.Key.Name == CbxPage.Text);

            if (page.Value.Columns.Contains(group))
            {
                page.Value.Columns.Remove(group);
            }

            if (page.Value.Rows.Contains(group))
            {
                page.Value.Rows.Remove(group);
            }


            var parent = treeListView1.GetParent(group) as Group;
            if (parent != null)
            {
                parent.Columns.Remove(group);
                if (parent.Columns.Count == 0)
                {
                    treeListView1.RemoveObject(parent);
                }
            }


            treeListView1.RemoveObject(group);
            page.Value.ReIndexItems();
            treeListView1.Refresh();
            CbxChangeRowsColumns();
        }

        private void treeListView1_MouseClick(object sender, MouseEventArgs e)
        {
            var group = treeListView1.SelectedObject as Group;
            if (group == null)
            {
                TbxNameElement.Clear();
                TbxDescElement.Clear();

            }

        }
        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (TbxNameElement.Text.Trim() == string.Empty && !_isEdit)
            {
                return;
            }

            var group = treeListView1.SelectedObject as Group;

            if (_isEdit)
            {
                if (!ValidElementField())
                {
                    return;
                }

                if (group != null)
                {
                    group.Description = TbxDescElement.Text.Trim();
                    group.Name = TbxNameElement.Text.Trim();
                }
                else
                {
                    var page = _report.Page.ElementAt(_currentPageIndex).Key;
                    page.Name = TbxNameElement.Text.Trim();
                    page.Description = TbxDescElement.Text.Trim();
                    int pageIndex = _currentPageIndex;
                    CbxPage.DataSource = _report.Page.Keys.ToList();
                    CbxPage.SelectedIndex = pageIndex;
                }

                TbxNameElement.Enabled = false;
                TbxDescElement.Enabled = false;
                BtnEditElement.Text = "Редактировать";
                _isEdit = false;
                treeListView1.Enabled = true;
                Menu.Enabled = true;
                CbxPage.Enabled = true;
                CbxShow.Enabled = true;


            }
            else
            {
                _isEdit = true;
                BtnEditElement.Text = "Сохранить";
                treeListView1.Enabled = false;
                TbxNameElement.Enabled = true;
                TbxDescElement.Enabled = true;
                Menu.Enabled = false;
                CbxPage.Enabled = false;
                CbxShow.Enabled = false;
            }



        }

        private void BtnCopyPage_Click(object sender, EventArgs e)
        {
            if (_report.Page.Count > 1)
            {
                using var form = new CopyDataForm(_report, _report.Page.ElementAt(_currentPageIndex).Value);
                form.ShowDialog();
            }

        }

        private void treeListView1_FormatRow(object sender, BrightIdeasSoftware.FormatRowEventArgs e)
        {
            var item = (Group)e.Model;
            if (item.IsGroup)
            {
                e.Item.Font = new Font(FontFamily.GenericSansSerif, 12, FontStyle.Italic);

            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}
