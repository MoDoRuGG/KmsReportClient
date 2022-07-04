using System;
using System.Windows.Forms;
using KmsReportClient.External;
using KmsReportClient.Global;

namespace KmsReportClient.Forms
{
    public partial class SettingsForm : Form
    {
        private readonly EndpointSoapClient _client;

        private bool _isEdit;
        private bool _isAddUser;

        public SettingsForm(EndpointSoapClient client)
        {
            InitializeComponent();
            CreateTree();

            this._client = client;
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            panelUser.Enabled = false;
            panelFilial.Enabled = false;

            tabControl.TabPages.Remove(tabFilial);
            tabControl.TabPages.Remove(tabUser);
        }

        private void CreateTree()
        {
            string[] root = {"Филиал", "Пользователь"};
            foreach (string element in root)
            {
                var tvRoot = new TreeNode {Text = element};
                treeSettings.Nodes.Add(tvRoot);
            }
        }

        private void TreeSettings_AfterSelect(object sender, TreeViewEventArgs e) =>
            SetTabPage();

        private void SetTabPage()
        {
            foreach (TabPage tp in tabControl.TabPages)
            {
                tabControl.TabPages.Remove(tp);
            }

            if (treeSettings.SelectedNode.Level == 0)
            {
                if (treeSettings.SelectedNode.Index == 0)
                {
                    tabControl.TabPages.Add(tabFilial);
                    ReadFilial();
                }
                else if (treeSettings.SelectedNode.Index == 1)
                {
                    tabControl.TabPages.Add(tabUser);
                    ReadUser();
                }
            }
        }

        private void ReadFilial()
        {
            txtbFilial.Text = CurrentUser.Filial;
            txtbFioHead.Text = CurrentUser.Director;
            txtbPlaceHead.Text = CurrentUser.DirectorPosition;
            txtbPhoneHead.Text = CurrentUser.DirectorPhone;

            panelFilial.Enabled = false;
            btnEditFilial.Visible = true;
        }

        private void ReadUser()
        {
            txtbFio.Text = CurrentUser.UserName;
            txtbPhone.Text = CurrentUser.Phone;
            txtbEmail.Text = CurrentUser.Email;

            btnAdd.Visible = true;
            btnEditUser.Visible = true;
            panelUser.Enabled = false;
        }

        private void BtnExit_Click(object sender, EventArgs e)
        {
            if (_isEdit || _isAddUser)
            {
                BtnExit.Text = "Закрыть";
                _isEdit = false;
                _isAddUser = false;

                if (treeSettings.SelectedNode.Text == "Филиал")
                {
                    ReadFilial();
                }
                else
                {
                    ReadUser();
                }
            }
            else
            {
                Close();
            }
        }

        private void ResetUserForm()
        {
            panelUser.Enabled = true;

            btnEditUser.Visible = false;
            btnAdd.Visible = false;
            BtnExit.Text = "Отмена";
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            txtbFio.Clear();
            txtbPhone.Clear();
            txtbEmail.Clear();
            ResetUserForm();

            _isAddUser = true;
        }

        private void btnEditUser_Click(object sender, EventArgs e)
        {
            panelUser.Enabled = true;
            ResetUserForm();

            _isEdit = true;
        }

        private void btnEditFilial_Click(object sender, EventArgs e)
        {
            _isEdit = true;
            panelFilial.Enabled = true;
            btnEditFilial.Visible = false;
            BtnExit.Text = "Отмена";
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (_isEdit || _isAddUser)
            {
                BtnExit.Text = "Закрыть";
                _isEdit = false;

                if (treeSettings.SelectedNode.Text == "Филиал")
                {
                    panelFilial.Enabled = false;
                    btnEditFilial.Visible = true;

                    string filialNaim = txtbFilial.Text;
                    string fio = txtbFioHead.Text;
                    string position = txtbPlaceHead.Text;
                    string phone = NormalizePhone(txtbPhoneHead.Text);

                    if (ValidFilial(filialNaim, fio, phone, position))
                    {
                        _client.EditFilial(CurrentUser.FilialCode, filialNaim, fio, position, phone);
                        MessageBox.Show("Информация о филиале успешно обновлена!", "Редактирование филиала!",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                }
                else
                {
                    string fio = txtbFio.Text;
                    string email = txtbEmail.Text;
                    string phone = NormalizePhone(txtbPhone.Text);

                    if (ValidUser(fio, phone, email))
                    {
                        _client.SaveUser(CurrentUser.FilialCode, fio, email, phone, _isAddUser, CurrentUser.IdUser);
                        MessageBox.Show("Пользователь успешно сохранен!", "Редактирование пользователя!",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    panelUser.Enabled = false;
                    if (_isAddUser)
                    {
                        ReadUser();
                        _isAddUser = false;
                    }
                    else
                    {
                        btnAdd.Visible = true;
                        btnEditUser.Visible = true;
                        panelUser.Enabled = false;
                    }
                }
            }
        }

        private string NormalizePhone(string phone)
        {
            return phone.Replace("(", "")
                .Replace(")", "")
                .Replace(" ", "")
                .Replace("_", "")
                .Replace("-", "");
        }

        private bool ValidUser(string fio, string phone, string email)
        {
            string message = "";
            if (string.IsNullOrEmpty(fio))
            {
                message += "Необходимо заполнить ФИО " + Environment.NewLine;
            }

            if (string.IsNullOrEmpty(email))
            {
                message += "Необходимо заполнить e-mail" + Environment.NewLine;
            }

            if (string.IsNullOrEmpty(phone))
            {
                message += "Необходимо заполнить телефон " + Environment.NewLine;
            }

            if (!string.IsNullOrEmpty(message))
            {
                MessageBox.Show(message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

        private bool ValidFilial(string filialNaim, string fio, string phone, string position)
        {
            string message = "";
            if (string.IsNullOrEmpty(filialNaim))
            {
                message += "Необходимо заполнить название филиала" + Environment.NewLine;
            }

            if (string.IsNullOrEmpty(fio))
            {
                message += "Необходимо заполнить ФИО руководителя" + Environment.NewLine;
            }

            if (string.IsNullOrEmpty(position))
            {
                message += "Необходимо заполнить должность руководителя" + Environment.NewLine;
            }

            if (string.IsNullOrEmpty(phone))
            {
                message += "Необходимо заполнить телефон руководителя" + Environment.NewLine;
            }

            if (!string.IsNullOrEmpty(message))
            {
                MessageBox.Show(message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }
    }
}