using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace KmsReportClient.Model.XML
{
   public class TemplateDynamicReport
    {
        [XmlElement("REPORT_NAME")] public string Name;
        [XmlElement("REPORT_DESCRIPTION")] public string ReportDescription;
        [XmlElement("REPORT_DATE")] public DateTime ReportDate;
        [XmlElement("IsUserRow")] public bool IsUserRow;
        [XmlArray("Emails")] public List<string> Executors;
        [XmlArray("TABLE")] public List<TemplateTableDynamic> tables;

       
    }

    public class TemplateTableDynamic
    {
        [XmlElement("TABLE_NAME")] public string Name;
        [XmlElement("TABLE_DESCRIPTION")] public string TableDescription;       
        [XmlArray("ROWS")] public List<TemplateRowDynamic> Rows;
        [XmlArray("COLUMNS")] public List<TemplateColumnDynamic> Columns;
    }

    [XmlRoot("ROW")]
    public class TemplateRowDynamic
    {
        [XmlElement("ROW_TEXT")] public string NameRow;
        [XmlElement("ROW_NUM")] public string IndexRow;
        [XmlElement("ROW_DESCRIPTION")] public string RowDescription;
        
    }

    [XmlRoot("COLUMN")]
    public class TemplateColumnDynamic
    {
        [XmlElement("COLUMN_TEXT")] public string NameColumn;
        [XmlElement("COLUMN_INDEX")] public string IndexColumn;
        [XmlElement("COLUMN_DESCRIPTION")] public string ColumnDescription;
        [XmlElement("COLUMN_DESCRIPTION_CHILD")] public List<TemplateColumnDynamic> ChildColumn;

    }
}
