using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KmsReportClient.External;
using KmsReportClient.Global;

namespace KmsReportClient.Forms
{
    public partial class EmailForm : Form
    {
        private readonly EndpointSoap _client;
        public EmailForm(EndpointSoap client)
        {
            InitializeComponent();
            _client = client;
            SetDgv();
        }


        public void SetCurrentRowForDgvAfterAdd()
        {
            int countRow = dgv.Rows.Count;
            if (countRow == 0)
            {
                return;
            }

            dgv.CurrentCell = dgv.Rows[countRow-1].Cells[1];
        }

        public void SetCurrentRowForDgvAfterEdit(int index)
        {
            int countRow = dgv.Rows.Count;
            if (countRow == 0)
            {
                return;
            }

            dgv.CurrentCell = dgv.Rows[index].Cells[1];
        }


        public void SetCurrentRowForDgvAfterDelete(int index)
        {
            int countRow = dgv.Rows.Count;
            if (countRow == 0 || countRow==1)
            {
                return;
            }

            if (index != 0)
            {
                dgv.CurrentCell = dgv.Rows[index - 1].Cells[1];
            } else
            {
                dgv.CurrentCell = dgv.Rows[index].Cells[1];
            }
           
        }

        private void SetDgv()
        {
            var emailList = _client.GetEmails(new GetEmailsRequest()).Body.GetEmailsResult;
            dgv.DataSource = emailList;
            dgv.Columns[0].Visible = false;
            dgv.Columns[1].HeaderText = "Email";
            dgv.Columns[2].HeaderText = "Описание";
        }

        

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            using var form = new EmailEditForm(_client);
            form.ShowDialog();
            SetDgv();
            SetCurrentRowForDgvAfterAdd();

        }

        private void BtnEdit_Click(object sender, EventArgs e)
        {
            if (dgv.CurrentRow == null)
            {
                return;
            }

            int index = dgv.CurrentRow.Index;
            int id = Convert.ToInt32(dgv.CurrentRow.Cells[0].Value);
            string email = dgv.CurrentRow.Cells[1].Value.ToString();
            string desc = dgv.CurrentRow.Cells[2].Value.ToString();
            using var form = new EmailEditForm(_client,id,email,desc);
            form.ShowDialog();
            SetDgv();
            SetCurrentRowForDgvAfterEdit(index);
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (dgv.CurrentRow == null)
            {
                return;
            }

            int id = Convert.ToInt32(dgv.CurrentRow.Cells[0].Value);
            if (MessageBox.Show("Вы действительно хотите удалить данную запись?","Удалить",MessageBoxButtons.YesNo,MessageBoxIcon.Question)==DialogResult.Yes)
            {
                _client.DeleteEmail(id);
            }
            int index = dgv.CurrentRow.Index;
            SetDgv();
            SetCurrentRowForDgvAfterDelete(index);



        }
    }
}
