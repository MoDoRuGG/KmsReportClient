using System.Collections.Generic;
using System.Xml.Serialization;

namespace KmsReportClient.Model.XML
{
    [XmlRoot("MAIL_TEMPLATE")]
    public class TemplateTextMail
    {
        [XmlElement("TEMPLATE")] public List<TemplateMail> templates;
    }

    [XmlRoot("TEMPLATE")]
    public class TemplateMail
    {
        [XmlElement("REPORT_TYPE")] public string ReportType;

        [XmlElement("TEXT_MAIL")] public string Text;
    }
}