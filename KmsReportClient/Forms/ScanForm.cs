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
using KmsReportClient.Report;
using KmsReportClient.Report.Basic;
using KmsReportClient.Service;
using NLog;

namespace KmsReportClient.Forms
{
    public partial class ScanForm : Form
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();
        private readonly FileProcessor _ftpProcessor = new FileProcessor();
        private TreeView _reportTree;
        IReportProcessor _processor;
        EndpointSoapClient _client;


        class ScanData
        {
            public int Num { get; set; }

            public string Name { get; set; }
        }

        public ScanForm(IReportProcessor processor, TreeView tree, EndpointSoapClient client)
        {
            InitializeComponent();
            this._processor = processor;
            this._client = client;
            this._reportTree = tree;


            if (CurrentUser.IsMain)
            {
                BtnLoad.Visible = false;
            }


            //var scans = _client.GetScans(processor.Report.IdFlow);

            //if (scans.Length != 0)
            //{
            //    dgvScan.DataSource = scans;
            //    dgvScan.Columns[0].Visible = false;
            //    dgvScan.Columns[1].HeaderText = "Файл";
            //    dgvScan.Columns[2].HeaderText = "Добавил";
            //    dgvScan.Columns[3].HeaderText = "Дата добавления";
            //    dgvScan.Columns[4].HeaderText = "Изменил";
            //    dgvScan.Columns[5].HeaderText = "Дата изменения";
            //    dgvScan.AutoResizeColumns();
            //}



            LbScan.DisplayMember = "Name";
            LbScan.ValueMember = "Num";

            if (!string.IsNullOrEmpty(processor.Report.Scan))
            {
                LbScan.Items.Add(new ScanData
                {
                    Num = 1,
                    Name = processor.Report.Scan
                });
            }

            if (!string.IsNullOrEmpty(processor.Report.Scan2))
            {
                LbScan.Items.Add(new ScanData
                {
                    Num = 2,
                    Name = processor.Report.Scan2
                });
            }

            if (!string.IsNullOrEmpty(processor.Report.Scan3))
            {
                LbScan.Items.Add(new ScanData
                {
                    Num = 3,
                    Name = processor.Report.Scan3
                });

            }

            if (!string.IsNullOrEmpty(processor.Report.Scan4))
            {
                LbScan.Items.Add(new ScanData
                {
                    Num = 4,
                    Name = processor.Report.Scan4
                });

            }

            if (!string.IsNullOrEmpty(processor.Report.Scan5))
            {
                LbScan.Items.Add(new ScanData
                {
                    Num = 5,
                    Name = processor.Report.Scan5
                });

            }

            if (!string.IsNullOrEmpty(processor.Report.Scan6))
            {
                LbScan.Items.Add(new ScanData
                {
                    Num = 6,
                    Name = processor.Report.Scan6
                });

            }

            if (!string.IsNullOrEmpty(processor.Report.Scan7))
            {
                LbScan.Items.Add(new ScanData
                {
                    Num = 7,
                    Name = processor.Report.Scan7
                });

            }

            if (!string.IsNullOrEmpty(processor.Report.Scan8))
            {
                LbScan.Items.Add(new ScanData
                {
                    Num = 8,
                    Name = processor.Report.Scan8
                });

            }

            if (!string.IsNullOrEmpty(processor.Report.Scan9))
            {
                LbScan.Items.Add(new ScanData
                {
                    Num = 9,
                    Name = processor.Report.Scan9
                });

            }

            if (!string.IsNullOrEmpty(processor.Report.Scan10))
            {
                LbScan.Items.Add(new ScanData
                {
                    Num = 10,
                    Name = processor.Report.Scan10
                });

            }

        }

        private void OpenScan(string uri)
        {
            try
            {
                string downloadFilename = _ftpProcessor
                    .DownloadFileFromWs(uri, "", _processor.FilialCode, _client);
                Process.Start(downloadFilename);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка открытия скана");
                MessageBox.Show("Ошибка открытия скана: " + ex.Message, "Ошибка", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public void UploadScan()
        {
            try
            {
                if (GlobalConst.SuccessStatuses.Contains(_processor.Report.Status) || _processor.Report.Status == ReportStatus.New)
                {
                    throw new Exception("Скан можно закачать только для отчетов, которые находится в статусах: " +
                        "'Сохранен в БД', 'Загружен скан', 'Отправлен на доработку'");
                }
                if (LbScan.Items.Count == 10 && LbScan.SelectedItem == null)
                {
                    throw new Exception("Максимальное количество сканов - 10 ");

                }

                var openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "PDF | *.pdf";
                if (openFileDialog1.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                int selectedIndex = 0;
                if (LbScan.SelectedItem != null)
                {
                    selectedIndex = LbScan.SelectedIndex;

                }
                string filename = openFileDialog1.FileName;
                var extension = Path.GetExtension(filename)?.ToLower() ?? "";

                if (extension != ".pdf")
                {
                    throw new Exception("Можно загрузить только документы с расширением PDF");
                }

                string savedFileName = "";
                if (LbScan.Items.Count >= 1)
                {
                    savedFileName = Path.GetFileName(openFileDialog1.FileName);
                }
                else
                {
                    savedFileName = GetFileName(extension);

                }

                _ftpProcessor.UploadFileToWs(filename, _processor.FilialCode, savedFileName, _client);

                int num = -1;
                if (LbScan.SelectedItem == null)
                {
                    int[] existsNums = LbScan.Items.Cast<ScanData>().Select(x => x.Num).Distinct().ToArray();

                    if (existsNums == null || existsNums.Length == 0)
                    {
                        num = 1;
                    }
                    else
                    {
                        for (int i = 1; i <= 10; i++)
                        {
                            if (!existsNums.Any(x => x == i))
                            {
                                num = i;
                                break;
                            }
                        }
                    }
                }
                else
                {
                    num = (LbScan.Items[selectedIndex] as ScanData).Num;
                }


                _processor.SaveScan(savedFileName, num);
                _processor.Report.Status = ReportStatus.Scan;
                _reportTree.SelectedNode.BackColor = GlobalConst.ColorScan;
                MessageBox.Show("Файл успешно загружен на сервер", "Загрузка завершена", MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                if (LbScan.SelectedItem != null)
                {
                    LbScan.Items[selectedIndex] = new ScanData
                    {
                        Num = num,
                        Name = savedFileName
                    };
                }
                else
                {
                    LbScan.Items.Add(new ScanData
                    {
                        Num = num,
                        Name = savedFileName
                    });
                }


                if (num == 1){_processor.Report.Scan = savedFileName;}
                if (num == 2) { _processor.Report.Scan2 = savedFileName; }
                if (num == 3) { _processor.Report.Scan3 = savedFileName; }
                if (num == 4) { _processor.Report.Scan4 = savedFileName; }
                if (num == 5) { _processor.Report.Scan5 = savedFileName; }
                if (num == 6) { _processor.Report.Scan6 = savedFileName; }
                if (num == 7) { _processor.Report.Scan7 = savedFileName; }
                if (num == 8) { _processor.Report.Scan8 = savedFileName; }
                if (num == 9) { _processor.Report.Scan9 = savedFileName; }
                if (num == 10) { _processor.Report.Scan10 = savedFileName; }


            }
            catch (Exception ex)
            {
                Log.Error(ex, $"Error saving scan of file");
                MessageBox.Show("Ошибка сохранения скана: " + ex.Message, "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void BtnOpen_Click(object sender, EventArgs e)
        {
            if (LbScan.SelectedItem != null)
                OpenScan((LbScan.SelectedItem as ScanData).Name);
        }
        private string GetFileName(string extension) =>
           $"{_processor.FilialCode}_{_processor.SmallName}_{_processor.Report.Yymm}{extension}";

        private void BtnLoad_Click(object sender, EventArgs e)
        {
            UploadScan();
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            DeleteScan();
        }

        private void DeleteScan()
        {

            if (LbScan.SelectedItem == null)
                return;

            if (MessageBox.Show("Вы действительно хотите удалить скан", "Удалить скан", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

          
            int num = (LbScan.SelectedItem as ScanData).Num;
            _processor.DeleteScan(num);

            if (num == 1) { _processor.Report.Scan = null; }
            if (num == 2) { _processor.Report.Scan2 = null; }
            if (num == 3) { _processor.Report.Scan3 = null; }
            if (num == 4) { _processor.Report.Scan4 = null; }
            if (num == 5) { _processor.Report.Scan5 = null; }
            if (num == 6) { _processor.Report.Scan6 = null; }
            if (num == 7) { _processor.Report.Scan7 = null; }
            if (num == 8) { _processor.Report.Scan8 = null; }
            if (num == 9) { _processor.Report.Scan9 = null; }
            if (num == 10) { _processor.Report.Scan10 = null; }

            if (String.IsNullOrEmpty(_processor.Report.Scan) && String.IsNullOrEmpty(_processor.Report.Scan2) && String.IsNullOrEmpty(_processor.Report.Scan3)
             && String.IsNullOrEmpty(_processor.Report.Scan4)
              && String.IsNullOrEmpty(_processor.Report.Scan5)
               && String.IsNullOrEmpty(_processor.Report.Scan6)
                && String.IsNullOrEmpty(_processor.Report.Scan7)
                 && String.IsNullOrEmpty(_processor.Report.Scan8)
                  && String.IsNullOrEmpty(_processor.Report.Scan9)
                   && String.IsNullOrEmpty(_processor.Report.Scan10)
                   )
            {
                _processor.Report.Status = ReportStatus.Saved;
            }

            LbScan.Items.RemoveAt(LbScan.SelectedIndex);

        }
    }
}
