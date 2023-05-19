using System.ComponentModel;

namespace KmsReportClient.Model.Enums
{
    public enum XmlFormTemplate
    {
        [Description("upd_ver2.xml")] UpdateXml,
        [Description("TemplateTextF262.xml")] F262,
        [Description("TemplateTextF294.xml")] F294,
        [Description("TemplateTextIizl.xml")] Iizl,
        [Description("TemplateTextPGQ.xml")] PgQ,
        [Description("TemplateTextZpzQ.xml")] ZpzQ,
        [Description("TemplateTextMail.xml")] TextMail,
        [Description("TemplateTextPG.xml")] Pg,
        [Description("TemplateTextZpz.xml")] Zpz,
        [Description("TemplateTextFOped.xml")] Oped,
        [Description("TemplateTextFOpedQ.xml")] OpedQ,
        [Description("TemplateTextFOpedU.xml")] OpedU,
        [Description("TemplateTextFCR.xml")] FCR,
        [Description("TemplateTextVac.xml")] Vac,
        [Description("TemplateFSSMonitoring.xml")] MFSS,
        [Description("TemplateMonitoringVCR.xml")] MVCR,
        [Description("TemplateTextProposal.xml")] Proposal,
        [Description("TemplateTextOpedFinance.xml")] OpedFinance,
        [Description("TemplateTextIizl2022.xml")] Iizl2022,
        [Description("TemplateTextCadre.xml")] Cadre,
        [Description("TemplateTextZpz10.xml")] Zpz10,
        [Description("TemplateTextEffectiveness.xml")] Effectiveness,
        [Description("TemplateTextZpzLethal.xml")] ZpzLethal
    }
}