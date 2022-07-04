using System.Collections.Generic;
using System.Xml.Serialization;

namespace KmsReportClient.Model.XML
{
    [XmlRoot("REPORT")]
    public class TemplateForm
    {
        [XmlElement("REPORT_NAME")] public string Name;

        [XmlElement("TABLE")] public List<TemplateTable> tables;
    }

    [XmlRoot("TABLE")]
    public class TemplateTable
    {
        [XmlElement("TABLE_NAME")] public string Name;

        [XmlElement("TABLE_DESCRIPTION")] public string TableDescription;
        [XmlElement("ROWS_COUNT")] public int RowsCount;
        [XmlElement("ROW")] public List<TemplateRow> Rows;
    }

    [XmlRoot("ROW")]
    public class TemplateRow
    {
        [XmlElement("ROW_TEXT")] public string Name;
        [XmlElement("ROW_NUM")] public string Num;
        [XmlElement("EXCLUSION")] public bool Exclusion;
        [XmlElement("EXCLUSION_CELL")] public string ExclusionCells;
    }
}