using BrightIdeasSoftware;
using KmsReportClient.Model;
using KmsReportClient.Model.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KmsReportClient.Forms
{
    public partial class CopyDataForm : Form
    {
        private DynamicReport _report;
        private ColumnRowReport _columnRowReportNew;
        Dictionary<PageReport, ColumnRowReport> oldPage;
        public CopyDataForm(DynamicReport report, ColumnRowReport columnRowReport)
        {
            InitializeComponent();
            _report = report;
            _columnRowReportNew = columnRowReport;
            if (_report == null)
                return;

            oldPage = _report.ClonePage();
            CbxPage.DataSource = _report.Page.Keys.ToList();
            CbxPage.DisplayMember = "Name";
            CbxPage.ValueMember = "Name";
            CbxPage.SelectedIndex = 0;
            CbxShow.DataSource = new List<string> { "Столбцы", "Строки" };
            CbxShow.SelectedIndex = 0;
            CreateTreeListView(TreeListOld);
            CreateTreeListView(TreeListNew);
            CbxChangeRowsColumns(TreeListOld);

        }

     
        private void CbxShow_SelectedIndexChanged(object sender, EventArgs e)
        {
            CbxChangeRowsColumns(TreeListOld);
            CbxChangeRowsColumnsNew();
        }

        public void CreateTreeListView(TreeListView treeListView)
        {
            treeListView.CanExpandGetter = model => ((Group)model).
                                                          Columns.Count > 0;
            treeListView.ChildrenGetter = delegate (object model)
            {
                return ((Group)model).Columns;
            };

            var IndexCol = new BrightIdeasSoftware.OLVColumn("Номер", "Index");
            IndexCol.AspectGetter = delegate (object x) { return ((Group)x).Index; };

            var NameCol = new BrightIdeasSoftware.OLVColumn("Наименование", "Name");
            NameCol.AspectGetter = delegate (object x) { return ((Group)x).Name; };

            treeListView.Columns.Add(IndexCol);
            treeListView.Columns.Add(NameCol);
            treeListView.AutoResizeColumns();


        }

        private void CbxPage_SelectedIndexChanged(object sender, EventArgs e)
        {
            CbxChangeRowsColumns(TreeListOld);
            CbxChangeRowsColumnsNew();
        }

        private void BtnAllNext_Click(object sender, EventArgs e)
        {

            var page = GetCurrentPage(CbxPage.SelectedIndex);
            if (CbxShow.SelectedIndex == 0)
            {
               
               // _columnRowReportNew.Columns = page.Value.Columns;
                foreach(var item in page.Value.Columns)
                {
                    _columnRowReportNew.Columns.Add(item.Clone(GetLastIndex(TableElement.Column)));
                }
                page.Value.Columns.Clear();
             
            }
            else
            {
                // _columnRowReportNew.Rows = page.Value.Rows;
                foreach (var item in page.Value.Columns)
                {
                    _columnRowReportNew.Rows.Add(item.Clone(GetLastIndex(TableElement.Row)));
                }
               

                page.Value.Rows.Clear();

            }
            TreeListOld.Roots = null;
            CbxChangeRowsColumnsNew();
        }
    

        public string GetLastIndex(TableElement element)
        {
            int result = 0;


            switch (element)
            {
                case TableElement.Row:
                    if (!_columnRowReportNew.Rows.Any())
                        return (result + 1).ToString();
                    result = _columnRowReportNew.Rows.Max(m => Convert.ToInt32(m.Index));
                    break;

                case TableElement.Column:
                    if (!_columnRowReportNew.Columns.Any())
                        return (result + 1).ToString();
                    int MaxIndex = 0;
                    int LostMaxIndex = 0;

                    foreach (var item in _columnRowReportNew.Columns)
                    {
                        MaxIndex = GetLastIndexInGroup(item);
                        if (MaxIndex > LostMaxIndex)
                            LostMaxIndex = MaxIndex;
                    }

                    result = LostMaxIndex;

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


        private void BtnNextOne_Click(object sender, EventArgs e)
        {
            if (TreeListOld.SelectedObject == null)
            {
                return;
            }
            var item = TreeListOld.SelectedObject as Group;
            var page = GetCurrentPage(CbxPage.SelectedIndex);


            if (CbxShow.SelectedIndex == 0)
            {
                page.Value.Columns.Remove(item);
                var copyItem = item.Clone(GetLastIndex(TableElement.Column));
                _columnRowReportNew.Columns.Add(copyItem);
            }
            else
            {
                page.Value.Rows.Remove(item);
                var copyItem = item.Clone(GetLastIndex(TableElement.Row));
                _columnRowReportNew.Rows.Add(copyItem);
            }


            CbxChangeRowsColumnsNew();
            CbxChangeRowsColumns(TreeListOld);


        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        public void CbxChangeRowsColumns(TreeListView treeListView)
        {
            if (CbxShow.SelectedItem is null)
                return;
            
            string CbxShowValue = CbxShow.SelectedItem.ToString();
            var ReportData = oldPage.ElementAt(CbxPage.SelectedIndex);
            ReportData.Value.ReIndexItems();
            if (CbxShowValue.Equals("Строки"))
            {

                treeListView.Roots = ReportData.Value.Rows;
            }
            else
            {
                treeListView.Roots = ReportData.Value.Columns;
            }

        }

        public void CbxChangeRowsColumnsNew()
        {
            _columnRowReportNew.ReIndexItems();
            if (CbxShow.SelectedIndex == 0)
            {
                TreeListNew.Roots = _columnRowReportNew.Columns;
            }
            else
            {
                TreeListNew.Roots = _columnRowReportNew.Rows;
            }

        }

        public KeyValuePair<PageReport, ColumnRowReport> GetCurrentPage(int pageIndex) => oldPage.ElementAt(pageIndex);

        private void BtnBackOne_Click(object sender, EventArgs e)
        {
            if (TreeListNew.SelectedObject == null)
            {
                return;
            }
            var item = TreeListNew.SelectedObject as Group;
            var page = GetCurrentPage(CbxPage.SelectedIndex);
           
            if (CbxShow.SelectedIndex == 0)
            {
                page.Value.Columns.Add(item);
                _columnRowReportNew.Columns.Remove(item);
            }
            else
            {
                page.Value.Rows.Add(item);
                _columnRowReportNew.Rows.Remove(item);
            }

            _columnRowReportNew.ReIndexItems();
            page.Value.ReIndexItems();

            CbxChangeRowsColumnsNew();
            CbxChangeRowsColumns(TreeListOld);
        }

        private void BtnBackAll_Click(object sender, EventArgs e)
        {
         
            var page = GetCurrentPage(CbxPage.SelectedIndex);
            if (CbxShow.SelectedIndex == 0)
            {
               // page.Value.Columns = _columnRowReportNew.Columns;           
                foreach(var item in _columnRowReportNew.Columns)
                {
                    page.Value.Columns.Add(item.Clone(GetLastIndex(TableElement.Column)));
                }
                _columnRowReportNew.Columns.Clear();
                
            }
            else
            {
               // page.Value.Rows = _columnRowReportNew.Columns;
                foreach (var item in _columnRowReportNew.Columns)
                {
                    page.Value.Rows.Add(item.Clone(GetLastIndex(TableElement.Row)));
                }
                _columnRowReportNew.Rows.Clear();
               
            }
            page.Value.ReIndexItems();
            CbxChangeRowsColumnsNew();
            CbxChangeRowsColumns(TreeListOld);
        }

        private void TreeListOld_FormatRow(object sender, FormatRowEventArgs e)
        {
            var item = (Group)e.Model;
            if (item.IsGroup)
            {
                e.Item.Font = new Font(FontFamily.GenericSansSerif, 12, FontStyle.Italic);

            }
        }

        private void TreeListNew_FormatRow(object sender, FormatRowEventArgs e)
        {
            var item = (Group)e.Model;
            if (item.IsGroup)
            {
                e.Item.Font = new Font(FontFamily.GenericSansSerif, 12, FontStyle.Italic);

            }
        }
    }
}
