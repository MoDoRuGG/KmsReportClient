using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Model;
using KmsReportClient.Report;
using KmsReportClient.Report.Basic;
using KmsReportClient.Service;

namespace KmsReportClient.Forms
{
    public partial class ScanDynamicForm : Form
    {

        private DynamicReportProcessor _processor;
        private TreeView _tree;
        private EndpointSoapClient _client;
        private ReportNodeTag _tag;
        private readonly FileProcessor _ftpProcessor = new FileProcessor();

        public ScanDynamicForm()
        {
            InitializeComponent();
        }

        public ScanDynamicForm(DynamicReportProcessor processor, ReportNodeTag tag, TreeView tree, EndpointSoapClient client) : this()
        {

            if (CurrentUser.IsMain)
            {
                BtnDelete.Visible = false;
                BtnLoad.Visible = false;
            }

            _processor = processor;
            _tree = tree;
            _client = client;
            _tag = tag;
            GetScan();
        }


        private void GetScan()
        {
            try
            {
                var scans = _client.GetScansDynamic(_tag.idFlow);
                //if (scans.Length > 0)
                //{
                    LbScan.DataSource = scans;
                    LbScan.ValueMember = "IdReportDynamicScan";
                    LbScan.DisplayMember = "FileName";
                //}

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnLoad_Click(object sender, EventArgs e)
        {
            try
            {
                var openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "PDF | *.pdf";
                if (openFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;


                }

                var flow = _processor.GetReportDynamicFlow(_tag.idFlow);
                _ftpProcessor.UploadFileToWs(openFileDialog1.FileName, flow.IdRegion, flow.IdRegion +"_" + Path.GetFileName(openFileDialog1.FileName), _client);
                _client.SaveDynamicScan(_tag.idFlow, flow.IdRegion + "_" + Path.GetFileName(openFileDialog1.FileName));
                MessageBox.Show("Файл успешно загружен на сервер", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                GetScan();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                if (LbScan.SelectedItem == null)
                    return;

                if (MessageBox.Show("Вы действительно хотите удалить скан?", "Вопрос", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                    return;

                var selItem = LbScan.SelectedItem as ReportDynamicScanModel;
                _client.DeleteDynamicScan(selItem.IdReportDynamicScan);
                GetScan();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void BtnOpen_Click(object sender, EventArgs e)
        {

            if (LbScan.SelectedItem == null)
                return;

            try
            {
                var selItem = LbScan.SelectedItem as ReportDynamicScanModel;
                var flow = _processor.GetReportDynamicFlow(_tag.idFlow);

                string downloadFilename = _ftpProcessor
                    .DownloadFileFromWs(selItem.FileName, "", flow.IdRegion, _client);
                Process.Start(downloadFilename);
            }
            catch (Exception ex)
            {
               
                MessageBox.Show("Ошибка открытия скана: " + ex.Message, "Ошибка", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
