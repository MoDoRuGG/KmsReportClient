using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.XML;

namespace KmsReportClient.Model
{
    public class DynamicReport
    {
        public int Id { get; set; }
        public string NameReport { get; set; }
        public DateTime DateReport { get; set; }

        public string DescriptionReport { get; set; }
        public bool IsUserRow { get; set; }
        public List<string> Executors = new List<string>();

        public Dictionary<PageReport, ColumnRowReport> Page { get; set; }

        public DynamicReport()
        {
            Page = new Dictionary<PageReport, ColumnRowReport>();
            Page.Add(new PageReport("1 вкладка", "1 вкладка"), new ColumnRowReport());

        }

        

        public Dictionary<PageReport, ColumnRowReport> ClonePage()
        {
            var result = new Dictionary<PageReport, ColumnRowReport>();
            foreach (var page in this.Page)
            {
                var theme = new PageReport
                {
                    Name = page.Key.Name,
                    Description = page.Key.Description,
                    Index = page.Key.Index
                };

                var columnRow = new ColumnRowReport
                {
                    Columns = page.Value.Columns.Select(x => new Group
                    {
                        Name = x.Name,
                        Index = x.Index,
                        Description = x.Description,
                        Columns = x.Columns.Select(y => new Group
                        {
                            Name = y.Name,
                            Index = y.Index,
                            Description = y.Description
                        }).ToList()

                    }).ToList(),

                    Rows = page.Value.Rows.Select(x => new Group
                    {
                        Name = x.Name,
                        Index = x.Index,
                        Description = x.Description
                    }).ToList()

                };

                result.Add(theme, columnRow);
            }
            return result;


        }

        public DynamicReport(TemplateDynamicReport Xml, int id)
        {
            this.Id = id;
            Page = new Dictionary<PageReport, ColumnRowReport>();
            Page.Add(new PageReport("1 вкладка", "1 вкладка"), new ColumnRowReport());
            ConvertToDynamicReport(Xml);



        }

        public void SetComboBox(ComboBox cmb)
        {
            var PageData = Page.Select(x => new { x.Key.Index, x.Key.Name, x.Key.Description });
            cmb.DataSource = PageData;
            cmb.DisplayMember = "Name";
            cmb.ValueMember = "Index";
        }

        private void ConvertToDynamicReport(TemplateDynamicReport Xml)
        {

           
            this.DateReport = Xml.ReportDate;
            this.NameReport = Xml.Name;
            this.DescriptionReport = Xml.ReportDescription;
            this.Executors = Xml.Executors;
            this.IsUserRow = Xml.IsUserRow;
            
            Dictionary<PageReport, ColumnRowReport> pages = new Dictionary<PageReport, ColumnRowReport>();

            //Переименовать Xml.Pages
            foreach (var page in Xml.tables)
            {
                var pageReport = new PageReport();
                pageReport.Name = page.Name;
                pageReport.Description = page.TableDescription;

                var CrW = new ColumnRowReport
                {
                    Columns = page.Columns.Select(x => new Group
                    {
                        Index = x.IndexColumn,
                        Name = x.NameColumn,
                        Description = x.ColumnDescription,
                        Columns = x.ChildColumn.Select(child => new Group
                        {
                            Index = child.IndexColumn,
                            Name = child.NameColumn,
                            Description = child.ColumnDescription
                        }).ToList()
                    }).ToList(),

                    Rows = page.Rows.Select(x => new Group
                    {
                        Index = x.IndexRow,
                        Name = x.NameRow,
                        Description = x.RowDescription
                    }).ToList()
                };

                pages.Add(pageReport, CrW);

            }

            this.Page = pages;

            //Добавить IsUserRow
            //this.IsUserRow= Xml.IsUserRow;
        }

    }

    public class PageReport : ElementBase
    {
        public PageReport(string name, string desc)
        {
            Name = name;
            Description = desc;
        }

        public PageReport()
        {

        }

    }
    public class ColumnRowReport
    {
        public List<Group> Columns = new List<Group>();
        public List<Group> Rows = new List<Group>();

        public ColumnRowReport()
        {
        }

        public void ReIndexItems()
        {
            int index = 1;
            foreach(var group in Columns.Where(x => x.IsGroup))
            {
                group.Index = string.Empty;
            }
            foreach(var col in Columns)
            {
                if (!col.IsGroup)
                {
                    col.Index = index.ToString();
                    index++;

                } else
                {
                    foreach(var subCol in col.Columns)
                    {
                        subCol.Index = index.ToString();
                        index++;
                    }
                }

                

            }

            index = 1;
            foreach (var row in Rows)
            {
                row.Index = index.ToString();

            }
        }

        public string GetLastIndex(TableElement element)
        {
            int result = 0;        
            switch (element)
            {
                case TableElement.Row:
                    if (Rows.Any())
                        return (result + 1).ToString();
                    result = Rows.Max(m => Convert.ToInt32(m.Index));
                    break;

                case TableElement.Column:
                    if (!Columns.Any())
                        return (result + 1).ToString();
                    int MaxIndex = 0;
                    int LostMaxIndex = 0;

                    foreach (var item in Columns)
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

        public static int GetLastIndexInGroup(Group group)
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



    }

    public class Group : ElementBase
    {

        public List<Group> Columns = new List<Group>();

        public bool IsGroup {
            get => this.Columns.Any();
        }

        public Group()
        {

        }

        public Group Clone(string index)
        {
            return new Group
            {
                Name = this.Name,
                Description = this.Description,
                Index = index,
                Columns = this.Columns.Select(x => new Group
                {
                    Name = x.Name,
                    Description = x.Description,
                    Index = x.Index
                }
                ).ToList()
            };

        }

        /// <summary>
        /// Если дочерний элемент  - true
        /// </summary>
        /// <param name="Parent Item"></param>
        /// <returns></returns>
        public static bool IsSubItem(object MainItem) => !(MainItem is null);


    }

    public class Column : ElementBase
    {

    }

    public class Row : ElementBase
    {

    }



}
