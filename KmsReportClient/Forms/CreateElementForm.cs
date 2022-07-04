
using KmsReportClient.Model;
using KmsReportClient.Model.Enums;
using KmsReportClient.Support;
using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace KmsReportClient.Forms
{
    public partial class CreateElementForm : Form
    {

        private ListView _listView;
        private Model.ColumnRowReport _columnRowReport;
        private TreeView _tree;
        private Group _group;
        private TreeNode _node;
        private DynamicReport _report;
        private TableElement _tableElement;
        private int _PageIndex;
        private ComboBox _CbxPage;
        public CreateElementForm()
        {
            InitializeComponent();
        }

        public CreateElementForm(ComboBox cbxPage, DynamicReport report, int pageIndex, TableElement tableElement)
        {
            InitializeComponent();
            _CbxPage = cbxPage;
            _report = report;
            _PageIndex = pageIndex;
            _tableElement = tableElement;
            this.Text = "Добавить: " + EnumUtils.GetDescription(_tableElement);
        }


        public CreateElementForm(object group, DynamicReport report, int pageIndex, TableElement tableElement)
        {
            InitializeComponent();

            _report = report;
            _PageIndex = pageIndex;
            _tableElement = tableElement;
            _group = group as Group;
            this.Text = "Добавить: " + EnumUtils.GetDescription(_tableElement);
        }


        public CreateElementForm(DynamicReport report, int pageIndex, TableElement tableElement)
        {
            InitializeComponent();

            _report = report;
            _PageIndex = pageIndex;
            _tableElement = tableElement;
            this.Text = "Добавить: " + EnumUtils.GetDescription(_tableElement);
        }

        public string GetLastIndex(TableElement element)
        {
            int result = 0;
            var data = _report.Page.ElementAt(_PageIndex).Value;

            switch (element)
            {
                case TableElement.Row:
                    if (!data.Rows.Any())
                        return (result + 1).ToString();
                    result = data.Rows.Max(m => Convert.ToInt32(m.Index));
                    break;

                case TableElement.Column:
                    if (!data.Columns.Any())
                        return (result + 1).ToString();
                    int MaxIndex = 0;
                    int LostMaxIndex = 0;

                    foreach (var item in data.Columns)
                    {
                        MaxIndex = GetLastIndexInGroup(item);
                        if (MaxIndex > LostMaxIndex)
                            LostMaxIndex = MaxIndex;
                    }

                    result = LostMaxIndex;
                    //result = _report.Page.ElementAt(_PageIndex).Value.Columns.Max(m => m.Index);
                    break;
            }

            return (result + 1).ToString();
        }

        public int GetLastIndexInGroup(Group group)
        {
            try
            {
                if (!group.IsGroup)
                    return Convert.ToInt32(group.Index);
                int result = group.Columns.Max(x => Convert.ToInt32(x.Index));
                return result;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                throw;
            }

        }

        private void ColumnAdd(String Name, String Desc)
        {

            if (_group is null)
            {
                _report.Page.ElementAt(_PageIndex).Value.Columns.Add(new Group
                {
                    Name = Name,
                    Description = Desc,               
                });
            }
            else
            {
                _report.Page.ElementAt(_PageIndex).Value.Columns.Where(c => c == _group).SingleOrDefault().Columns.
                    Add(new Group
                    {
                        Name = Name,
                        Description = Desc,
              
                    });

            }

          

        }

        private void RowAdd(String Name, String Desc)
        {

            _report.Page.ElementAt(_PageIndex).Value.Rows.Add(new Model.Group
            {
                Name = Name,
                Description = Desc,           
            });

        }

        private void GroupAdd(String Name, String Desc)
        {
            _report.Page.ElementAt(_PageIndex).Value.Columns.Add(new Group
            {
                Name = Name,
                Description = Desc,
                Index = "",
                Columns = new System.Collections.Generic.List<Group>()
                {
                    new Group
                    {
                       //Index= /GetLastIndex(TableElement.Column),
                       Name = "New column",
                       Description = "Description"

                    }
                }

            });

        }

        private void PageAdd(String Name, String Desc)
        {

            _columnRowReport = new ColumnRowReport();

            DialogResult result = MessageBox.Show("Скопировать строки и столбцы из других вкладок?", "Новая вкладка", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                using var copyDataForm = new CopyDataForm(_report, _columnRowReport);
                copyDataForm.ShowDialog();
            }

            _report.Page.Add(new Model.PageReport(Name, Desc), _columnRowReport);
            _CbxPage.DataSource = _report.Page.Keys.ToList();
            Close();

        }

        private bool ValidateField()
        {
            if (TbxName.Text.Trim() == String.Empty)
            {
                MessageBox.Show("Введите наименование", "Ошибка создания", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void TbxClear()
        {
            TbxDesc.Clear();
            TbxName.Clear();
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (!ValidateField())
            {
                return;
            }

            string Name = TbxName.Text.Trim();
            string Desc = TbxDesc.Text.Trim();

            string NodeText;
            if (_node == null)
            {
                NodeText = "all";

            }
            else
            {
                NodeText = _node.Text;
            }


            if (NodeText == null)
                NodeText = "all";

            switch (_tableElement)
            {
                case TableElement.Column:
                    ColumnAdd(Name, Desc);
                    _report.Page.ElementAt(_PageIndex).Value.ReIndexItems();
                    break;
                case TableElement.Row:
                    RowAdd(Name, Desc);
                    _report.Page.ElementAt(_PageIndex).Value.ReIndexItems();
                    break;
                case TableElement.Page:
                    PageAdd(Name, Desc);
                    this.Close();
                    break;
                case TableElement.Group:
                    GroupAdd(Name, Desc);
                    _report.Page.ElementAt(_PageIndex).Value.ReIndexItems();
                    this.Close();
                    break;
            }
            TbxClear();
        }

        private void BtnCancel_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
