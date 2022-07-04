using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Xml.Serialization;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Model;
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.XML;
using KmsReportClient.Support;
using NLog;

namespace KmsReportClient.Forms
{
    public partial class NotificationForm : Form
    {
        private const string TemplatesFolder = "Template\\";

        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        private bool IsQuery;

        private readonly EndpointSoapClient _client;
        private readonly List<KmsReportDictionary> _regions;
        private readonly TemplateTextMail _templateText;
        private readonly DynamicReport _dynamicReport;
        private List<KmsReportDictionary> _emails;


        public NotificationForm(
            EndpointSoapClient client,
            List<KmsReportDictionary> regions,
            List<KmsReportDictionary> reportDictionary)
        {
            InitializeComponent();
            _client = client;
            SetBaseReportInterface(regions, reportDictionary);
            _templateText = ReadTemplateXml(TemplatesFolder + XmlFormTemplate.TextMail.GetDescription());
            _regions = regions;
        }

        public NotificationForm(EndpointSoapClient client, DynamicReport report, List<string> currentEmails)
        {
            InitializeComponent();

            IsQuery = true;
            _client = client;
            _emails = _client.GetEmails().ToList();
            _dynamicReport = report;
            var currentEmailAdress = _emails.Where(x => currentEmails.Contains(x.Value)).Select(x => x.ForeignKey).ToList();
            SetDynamicReportInterface(currentEmailAdress);


        }

        public void SetBaseReportInterface(List<KmsReportDictionary> regions, List<KmsReportDictionary> reportDictionary)
        {
            chkbDefault.Checked = true;
            panel.Enabled = false;
            txtbTheme.Enabled = false;

            CmbFilials.DisplayMember = "Value";
            CmbFilials.ValueMember = "Key";
            CmbFilials.DataSource = regions;

            CmbReport.DataSource = reportDictionary.ToList();
            CmbReport.ValueMember = "Key";
            CmbReport.DisplayMember = "Value";
            CmbMonth.DataSource = GlobalConst.Months;

            CmbMonth.SelectedIndex = DateTime.Today.Month - 1;
            NudYear.Value = DateTime.Today.Year;
            RbtnLater.Checked = true;




            SetTextTemplate();

        }

        public void SetDynamicReportInterface(List<string> currentEmails)
        {


            CmbReport.Visible = false;
            CmbMonth.Visible = false;
            NudYear.Visible = false;

            RbtnLater.Visible = false;
            RbtnRefuse.Visible = false;
            RbtnSelect.Visible = false;

            chkbDefault.Visible = false;
            BtnSaveText.Visible = false;

            panel.Location = new System.Drawing.Point(4, 12);
            panel.Size = new System.Drawing.Size(461, 168);
            TxtbListFilial.Size = new System.Drawing.Size(442, 120);

            for (int i = 0; i < currentEmails.Count; i++)
            {
                if (i != currentEmails.Count - 1)
                {
                    TxtbListFilial.Text += currentEmails[i] + ",";

                }
                else
                {
                    TxtbListFilial.Text += currentEmails[i];

                }
            }

            TxtbListFilial.Text.Replace(", ", "");

            CmbFilials.DataSource = _emails;
            CmbFilials.ValueMember = "Key";
            CmbFilials.DisplayMember = "ForeignKey";

            txtbTheme.Text = String.Format($"Новый запрос от {_dynamicReport.DateReport.ToShortDateString()} - {_dynamicReport.NameReport}");
            TxtbText.Text = _dynamicReport.DescriptionReport;

        }

        private TemplateTextMail ReadTemplateXml(string xmlPath)
        {
            try
            {
                var xmlDoc = XDocument.Load(xmlPath);
                var xmlSerializer = new XmlSerializer(typeof(TemplateTextMail));
                if (xmlDoc.Root != null)
                {
                    using var reader = xmlDoc.Root.CreateReader();
                    return (TemplateTextMail)xmlSerializer.Deserialize(reader);
                }
            }
            catch (Exception ex)
            {
                Log.Error($"Форма уведомлений. Ошибка получения данных из шаблона: {ex}");
                MessageBox.Show($"Ошибка получения данных из шаблона: {ex}", "Ошибка", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }

            return null;
        }

        private void SaveTemplateText()
        {
            try
            {
                var temp = _templateText.templates.SingleOrDefault(t => t.ReportType == CmbReport.Text);
                if (temp != null)
                {
                    temp.Text = TxtbText.Text;
                    string filename = AppDomain.CurrentDomain.BaseDirectory +
                        "Template\\" + XmlFormTemplate.TextMail.GetDescription();

                    using var writer = new FileStream(filename, FileMode.Create);
                    var xmlSerializer = new XmlSerializer(typeof(TemplateTextMail));
                    xmlSerializer.Serialize(writer, _templateText);
                }
            }
            catch (Exception ex)
            {
                Log.Error($"Форма уведомлений. Ошибка сохранения данных в шаблон: {ex}");
                MessageBox.Show($"Ошибка сохранения данных в шаблон: {ex}", "Ошибка", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void SetTextTemplate()
        {
            if (_templateText == null)
            {
                return;
            }

            var template = _templateText.templates.SingleOrDefault(t => t.ReportType == CmbReport.Text)?.Text ?? "";
            TxtbText.Text = template;
        }

        private void SetVisiblePanel() => panel.Enabled = RbtnSelect.Checked;

        private bool IsNotValidForm()
        {
            var message = "";
            if (RbtnSelect.Checked && TxtbListFilial.Text.Length == 0)
                message = "Необходимо выбрать филиалы для отправки, так как выбран пункт 'ВЫбрать из списка' ";
            if (!chkbDefault.Checked && txtbTheme.Text.Length == 0)
                message = Environment.NewLine + "Необходимо заполнить тему письма";
            if (TxtbText.Text.Length == 0)
                message = Environment.NewLine + "Необходимо заполнить текст письма";
            if (message.Length > 0)
            {
                MessageBox.Show(message, "Ошибка заполнения формы", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return true;
            }
            return false;
        }

        private bool IsNotValidFormDynamic()
        {
            var message = "";
            if (TxtbListFilial.Text.Length == 0)
            {
                message = "Необходимо выбрать адресатов!";
            }

            if (TxtbText.Text.Length == 0)
            {
                message = Environment.NewLine + "Необходимо заполнить текст письма";
            }

            if (message.Length > 0)
            {
                MessageBox.Show(message, "Ошибка заполнения формы", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return true;
            }
            return false;
        }

        private void SendMessage()
        {
            if (IsNotValidForm())
                return;

            try
            {
                Log.Info("Старт отправки уведомлений");

                var reportType = CmbReport.SelectedValue.ToString();
                var isRefuse = RbtnRefuse.Checked;
                var yymm = YymmUtils.GetYymmFromInt(NudYear.Value, Convert.ToInt32(CmbMonth.SelectedIndex + 1));
                var text = TxtbText.Text;
                var theme = !chkbDefault.Checked
                    ? txtbTheme.Text
                    : $"{CmbReport.Text} за {CmbMonth.Text} {NudYear.Value} года. Необходимо сдать отчет!";


                var request = new NotificationRequest
                {
                    CurrentEmail = CurrentUser.Email,
                    Filials = CreateFilialsArray(),
                    IsRefuse = isRefuse,
                    Text = text,
                    Theme = theme,
                    Yymm = yymm,
                    ReportType = reportType
                };
                var result = _client.SendNotification(request);
                MessageBox.Show(result, "Отправка уведомлений завершена!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                Log.Info("Успешное завершение отправки уведомлений");
            }
            catch (Exception ex)
            {
                Log.Error($"Форма уведомлений. Ошибка отправки электронной почты: {ex}");
                MessageBox.Show($"Ошибка отправки электронной почты: {ex}", "Ошибка", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void SendMessageDynamic()
        {
            if (IsNotValidFormDynamic())
                return;

            try
            {
                Log.Info("Старт отправки уведомлений");

                var theme = txtbTheme.Text.Trim();
                var body = TxtbText.Text.Trim();
                var emails = CreateEmailArray();

                _client.SendEmail(emails, theme, body);
                MessageBox.Show("Отправка уведомлений завершена!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);

                Log.Info("Успешное завершение отправки уведомлений");
            }
            catch (Exception ex)
            {
                Log.Error($"Форма уведомлений. Ошибка отправки электронной почты: {ex}");
                MessageBox.Show($"Ошибка отправки электронной почты: {ex}", "Ошибка", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private ArrayOfString CreateFilialsArray()
        {
            var filials = TxtbListFilial.Text.Split(',').Select(x => x.Trim()).ToArray();
            var filialCodes = _regions.Where(x => filials.Contains(x.Value)).Select(x => x.Key);

            var arrayOfFilial = new ArrayOfString();
            arrayOfFilial.AddRange(filialCodes);
            return arrayOfFilial;
        }


        private ArrayOfString CreateEmailArray()
        {
            var emailDesc = TxtbListFilial.Text.Split(',').Select(x => x.Trim()).ToArray();
            var emailslist = _emails.Where(x => emailDesc.Contains(x.ForeignKey)).Select(x => x.Value);

            var arrayOfFilial = new ArrayOfString();
            arrayOfFilial.AddRange(emailslist);
            return arrayOfFilial;
        }

        private void ClearForm()
        {

            if (!IsQuery)
            {
                CmbMonth.SelectedIndex = DateTime.Today.Month - 1;
                CmbReport.SelectedIndex = 0;
                NudYear.Value = DateTime.Today.Year;
                RbtnLater.Checked = true;
            }
            else
            {
                CmbFilials.SelectedIndex = 0;
                TxtbListFilial.Clear();
            }
            TxtbText.Clear();
            txtbTheme.Clear();


        }

        private void ChkbDefault_CheckedChanged(object sender, EventArgs e) => txtbTheme.Enabled = !chkbDefault.Checked;

        private void BtnPlus_Click(object sender, EventArgs e) =>
            GlobalUtils.AddValueInTextBox(CmbFilials, TxtbListFilial, true, false);

        private void BtnMinus_Click(object sender, EventArgs e) => GlobalUtils.DeleteValueFromTextBox(TxtbListFilial);

        private void RbtnLater_CheckedChanged(object sender, EventArgs e) => SetVisiblePanel();

        private void RbtnRefuse_CheckedChanged(object sender, EventArgs e) => SetVisiblePanel();

        private void RbtnSelect_CheckedChanged(object sender, EventArgs e) => SetVisiblePanel();

        private void BtnSaveText_Click(object sender, EventArgs e) => SaveTemplateText();

        private void CmbReport_SelectedIndexChanged(object sender, EventArgs e) => SetTextTemplate();

        private void BtnSend_Click(object sender, EventArgs e)
        {
            if (IsQuery)
            {

                SendMessageDynamic();
            }
            else
            {
                SendMessage();

            }
        }

        private void BtnClear_Click(object sender, EventArgs e) => ClearForm();
    }
}