﻿using System.ComponentModel;

namespace KmsReportClient.Model.Enums
{
    public enum XmlFormTemplate
    {
        [Description("upd_ver2.xml")] UpdateXml,
        [Description("TemplateTextF262.xml")] F262,
        [Description("TemplateTextF294.xml")] F294,
        [Description("TemplateTextIizl.xml")] Iizl,
        [Description("TemplateTextPGQ.xml")] PgQ,
        [Description("TemplateTextZpzQ2025.xml")] ZpzQ2025,
        [Description("TemplateTextZpzQ.xml")] ZpzQ,
        [Description("TemplateTextMail.xml")] TextMail,
        [Description("TemplateTextPG.xml")] Pg,
        [Description("TemplateTextZpz.xml")] Zpz,
        [Description("TemplateTextZpz2025.xml")] Zpz2025,
        [Description("TemplateTextFOped.xml")] Oped,
        [Description("TemplateTextFOpedQ.xml")] OpedQ,
        [Description("TemplateTextFOpedU.xml")] OpedU,
        [Description("TemplateTextFCR.xml")] FCR,
        [Description("TemplateTextVac.xml")] Vac,
        [Description("TemplateFSSMonitoring.xml")] MFSS,
        [Description("TemplateMonitoringVCR.xml")] MVCR,
        [Description("TemplateTextProposal.xml")] Proposal,
        [Description("TemplateTextOpedFinance.xml")] OpedFinance,
        [Description("TemplateTextOpedFinance3.xml")] OpedFinance3,
        [Description("TemplateTextIizl2022.xml")] Iizl2022,
        [Description("TemplateTextCadre.xml")] Cadre,
        [Description("TemplateTextZpz10.xml")] Zpz10,
        [Description("TemplateTextZpz10_2025.xml")] Zpz10_2025,
        [Description("TemplateTextEffectiveness.xml")] Effectiveness,
        [Description("TemplateTextZpzLethal.xml")] ZpzLethal,
        [Description("TemplateTextZpz2025Lethal.xml")] Zpz2025Lethal,
        [Description("TemplateTextReqVCR.xml")] ReqVCR,
        [Description("TemplateTextQuantity.xml")] Quantity,
        [Description("TemplateTextTargetedAllowances.xml")] TarAllow,
        [Description("TemplateTextPVPLoad.xml")] PVPL,
        [Description("TemplateTextDoff.xml")] Doff,
    }
}