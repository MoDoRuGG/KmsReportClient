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
using KmsReportClient.Spravochnik;

namespace KmsReportClient.Forms
{
    public partial class EmailEditForm : Form
    {

        private readonly EmailProcessor _emailProcessor;
        private bool isEdit;
        private int _id;

        public EmailEditForm(EndpointSoap client)
        {
            InitializeComponent();
            this.Text = "Добавить";
            _emailProcessor = new EmailProcessor(client);

        }

        public EmailEditForm(EndpointSoap client, int id, string email, string desc)
        {
            InitializeComponent();
            this.Text = "Изменить";
            _emailProcessor = new EmailProcessor(client);
            isEdit = true;
            tbxDesc.Text = desc;
            tbxEmail.Text = email;
            _id = id;

        }



        private void BtnSave_Click(object sender, EventArgs e)
        {
            string desc = tbxDesc.Text.Trim();
            string email = tbxEmail.Text.Trim();
            if (!ValidField(desc, email))
            {
                return;
            }

            if (!isEdit)
            {
                _emailProcessor.AddEmail(email, desc);
            }
            else
            {
                _emailProcessor.EditEmail(_id, email, desc);

            }



            this.Close();



        }

        public bool ValidField(string desc, string email)
        {

            if (desc == String.Empty || email == string.Empty)
            {
                MessageBox.Show("Остались незаполненные поля.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (email.Length < 6)
            {
                MessageBox.Show("Слишком короткий Email", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;

        }
    }
}
