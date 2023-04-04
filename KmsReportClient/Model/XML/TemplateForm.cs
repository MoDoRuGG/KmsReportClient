using System.Collections.Generic;
using System.Xml.Serialization;

namespace KmsReportClient.Model.XML
{
    [XmlRoot("REPORT")]
    public class TemplateForm
    {
        [XmlElement("REPORT_NAME")] public string ReportName_fromxml;

        [XmlElement("TABLE")] public List<TemplateTable> Tables_fromxml;
    }

    [XmlRoot("TABLE")]
    public class TemplateTable
    {
        [XmlElement("TABLE_NAME")] public string TableName_fromxml;

        [XmlElement("TABLE_DESCRIPTION")] public string TableDescription_fromxml;
        [XmlElement("ROWS_COUNT")] public int RowsCount_fromxml;
        [XmlElement("ROW")] public List<TemplateRow> Rows_fromxml;
    }

    [XmlRoot("ROW")]
    public class TemplateRow
    {
        [XmlElement("ROW_TEXT")] public string RowText_fromxml;
        [XmlElement("ROW_NUM")] public string RowNum_fromxml;
        [XmlElement("EXCLUSION")] public bool Exclusion_fromxml;
        [XmlElement("EXCLUSION_CELL")] public string ExclusionCells_fromxml;
    }
}