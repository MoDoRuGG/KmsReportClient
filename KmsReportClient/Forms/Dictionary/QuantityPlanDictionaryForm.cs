using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Support;

namespace KmsReportClient.Forms.Dictionary
{
    public partial class QuantityPlanDictionaryForm : Form
    {

        private readonly EndpointSoap _client;
        public QuantityPlanDictionaryForm(EndpointSoap client)
        {
            InitializeComponent();
            _client = client;
            nudYear.Value = DateTime.Now.Year;

            CreateDataGridView();
            FillDataGridView();



        }



        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void CreateDataGridView()
        {
            dgvDictionary.Columns.Clear();
            string year = GetYear().Substring(2);

            dgvDictionary.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "Филиал",
                Name = "filialName",
                ReadOnly = true,
                DefaultCellStyle = new DataGridViewCellStyle() { BackColor = Color.LightGray }

            });

            //dgvDictionary.Columns.Add(new DataGridViewTextBoxColumn
            //{
            //    HeaderText = "Год",
            //    Name = "YearValue",
            //    ReadOnly = true,
            //    DefaultCellStyle = new DataGridViewCellStyle() { BackColor = Color.LightGray }

            //});

            //Console.WriteLine(CultureInfo.CurrentCulture.Name);
            foreach (var month in GlobalConst.MonthsWithNumber)
            {
                dgvDictionary.Columns.Add(new DataGridViewTextBoxColumn
                {
                    HeaderText = month.Value,
                    Name = year + month.Key,
                    DefaultCellStyle = new DataGridViewCellStyle() { Format = "#,#" }

                });
            }

            foreach (var region in CurrentUser.Regions.Where(x => x.Key != "RU" && x.Key != "RU-KHA"))
            {

                int rowIndex = dgvDictionary.Rows.Add();
                dgvDictionary.Rows[rowIndex].Tag = region.Key;
                dgvDictionary.Rows[rowIndex].Cells["filialName"].Value = region.Value;

            }


        }

        private void FillDataGridView()
        {
            try
            {
                string year = GetYear();
                var data = _client.GetQuantityPlanList(new GetQuantityPlanListRequest
                {
                    Body = new GetQuantityPlanListRequestBody
                    {
                        year = year
                    }
                }).Body.GetQuantityPlanListResult;

                if (data != null)
                {
                    foreach (DataGridViewRow row in dgvDictionary.Rows)
                    {
                        foreach (DataGridViewColumn column in dgvDictionary.Columns)
                        {
                            var monthPlanData = data.FirstOrDefault(x => x.Yymm == column.Name && x.IdRegion == row.Tag.ToString());
                            if (monthPlanData != null)
                            {
                                row.Cells[monthPlanData.Yymm].Value = monthPlanData.Value;
                            }
                        }
                    }


                    //for (int i = 0; i < dgvDictionary.Rows.Count; i++)
                    //{
                    //    SetColumnYearValueForRow(i);
                    //}

                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка получения списка справочника", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private string GetYear() => nudYear.Value.ToString();

        private void btnSave_Click(object sender, EventArgs e)
        {
            Save();
        }

        private void Save()
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                List<QuantityPlanDictionaryItem> requestData = new List<QuantityPlanDictionaryItem>();
                foreach (DataGridViewRow row in dgvDictionary.Rows)
                {
                    foreach (DataGridViewColumn column in dgvDictionary.Columns)
                    {
                        if (column.Name == "filialName" || column.Name == "YearValue")
                            continue;

                        var value = row.Cells[column.Name].Value == null ? "" : row.Cells[column.Name].Value.ToString().Replace(" ", "");
                        decimal res = 0.00m;

                        requestData.Add(new QuantityPlanDictionaryItem
                        {
                            IdRegion = row.Tag.ToString(),
                            Value = decimal.TryParse(value, out res) ? res : 0.0m,
                            Yymm = column.Name
                        });

                    }

                }

                _client.SaveQuantityPlanList(new SaveQuantityPlanListRequest { Body = new SaveQuantityPlanListRequestBody { plans = requestData.ToArray() } });

                MessageBox.Show("Сохранено!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                FillDataGridView();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка сохранения", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void SetColumnYearValueForRow(int rowIndex)
        {
            decimal res = 0.00m;
            dgvDictionary.Rows[rowIndex].Cells["YearValue"].Value = dgvDictionary.Rows[rowIndex].Cells
                .Cast<DataGridViewTextBoxCell>()
                .Where(x => x.ColumnIndex != 0 && x.ColumnIndex != 1)
                .Sum(x => decimal.TryParse(x.Value == null ? "" : x.Value.ToString(), out res) ? res : 0.00m);
        }

        private void dgvDictionary_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {

        }

        private void dgvDictionary_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            var dec = 0.00m;
            string value = dgvDictionary.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == null ? "" : dgvDictionary.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            if (decimal.TryParse(value, out dec))
            {
                dgvDictionary.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = dec.ToString("#,#");

            }

        }

        private void nudYear_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                Cursor = Cursors.WaitCursor;
                dgvDictionary.Rows.Clear();
                dgvDictionary.Columns.Clear();
                CreateDataGridView();
                FillDataGridView();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void dgvDictionary_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            //Console.WriteLine(dgvDictionary.Rows[e.RowIndex].Cells[e.ColumnIndex].Value);
        }

        private void dgvDictionary_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {


        }

        private void dgvDictionary_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //DataGridViewTextBoxEditingControl ctr = e.Control as DataGridViewTextBoxEditingControl;

            //if (ctr != null)
            //{
            //    ctr.TextChanged += delegate
            //    {
            //        decimal dec = 0.00m;
            //        Console.WriteLine(e.Control.Text);
            //        if (decimal.TryParse(e.Control.Text, out dec))
            //        {
            //            dgvDictionary.Rows[dgvDictionary.CurrentCell.RowIndex].Cells[dgvDictionary.CurrentCell.ColumnIndex].Value = dec.ToString("#,#");
            //        }
            //        bool textChanged = true;
            //    };
            //}
        }

        private void dgvDictionary_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void PasteFromClipBoard()
        {
            if (dgvDictionary.CurrentCell == null)
                return;

            string clipboardText = Clipboard.GetText();
            if (string.IsNullOrEmpty(clipboardText))
                return;

            dgvDictionary.CurrentCell.Value = clipboardText;


        }

        protected override bool ProcessCmdKey(ref System.Windows.Forms.Message msg, System.Windows.Forms.Keys keyData) // И переопределяем метод обработки нажатия управляющих клавиш, унаследованный от Control
        {
            if ((keyData == (Keys.V | Keys.Control)))  // Если нажаты CTRL+D
            {
                PasteFromClipBoard();
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void dgvDictionary_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //SetColumnYearValueForRow(e.RowIndex);
        }
    }
}
