using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Forms;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Properties;
using NLog;

namespace KmsReportClient.Forms
{
    public partial class AuthorizationForm : Form
    {
        public static bool Status;
        
        private const string SavedSettings = "Serializable\\currentUser.dat";
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        private EndpointSoapClient _client;
        private ClientContext _context;

        public AuthorizationForm()
        {
            InitializeComponent();
            Directory.CreateDirectory("Serializable");
            Directory.CreateDirectory("Temp");
            CurrentUser.IsMain = Settings.Default.IsHeadOffice == "Y";
        }

        private void Authorization_Load(object sender, EventArgs e)
        {
            try
            {
                _client = new EndpointSoapClient();
                _context = _client.CollectClientContext();
                
                                
                FillComboBoxes(_context.Regions);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка получения сохраненных данных для входа");
                MessageBox.Show($"Сервер временно недоступен {ex.Message}", "Сервер временно недоступен", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
        }

        private void Btn_enter_Click(object sender, EventArgs e)
        {
            var idUser = CmbLogin.SelectedValue.ToString();
            var password = TxtbPassword.Text;

            try
            {
                if (CheckPassword(idUser, password))
                {
                    Visible = false;
                    Status = true;

                    SaveEnterData();
                    Close();
                }
                else
                {
                    MessageBox.Show(@"Неверный пароль!", @"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка проверки пароля");
                MessageBox.Show($"Ошибка проверки пароля: {ex}", @"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Autorization_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Enter:
                    BtnEnter.PerformClick();
                    break;
                case Keys.Escape:
                    BtnExit.PerformClick();
                    break;
            }
        }

        private void Btn_exit_Click(object sender, EventArgs e) =>
            Application.Exit();

        private void CmbFilial_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_context.Users!= null)
            {
                CmbLogin.DataSource = _context.Users.Where(
                    x => x.ForeignKey == CmbFilial.SelectedValue.ToString()).ToList();
            }
        }

        private void FillComboBoxes(KmsReportDictionary[] regions)
        {
            List<KmsReportDictionary> regionsDs = !CurrentUser.IsMain
                    ? regions.Where(x => x.Key != "RU").ToList()
                    : regions.Where(x => x.Key == "RU").ToList();

            CmbFilial.DisplayMember = "Value";
            CmbFilial.ValueMember = "Key";
            CmbFilial.DataSource = regionsDs;

            CmbLogin.DisplayMember = "Value";
            CmbLogin.ValueMember = "Key";

            string savedData = CheckSavedData();
            string filialCode = regionsDs[0].Key;

            string[] parts = savedData.Split('@');
            if (parts.Length == 2)
            {
                filialCode = regionsDs.SingleOrDefault(x => x.Value == parts[0])?.Key;
                if (filialCode == null)
                {
                    filialCode = regionsDs[0].Key;
                }
                CmbLogin.DataSource = GetCurrentUserList(filialCode);

                CmbFilial.Text = parts[0];
                CmbLogin.Text = parts[1];                
            }
            else
            {
                CmbLogin.DataSource = GetCurrentUserList(filialCode);
            }
            
        }

        private void SaveEnterData()
        {
            string savedLogin = CmbFilial.Text + "@" + CmbLogin.Text;
            var binFormat = new BinaryFormatter();
            using var fStream = new FileStream(SavedSettings, FileMode.Create, FileAccess.Write, FileShare.None);
            binFormat.Serialize(fStream, savedLogin);
        }

        private bool CheckPassword(string idUser, string password)
        {
            var employee = _client.CheckPasswordNew(idUser, password);
            if (employee != null)
            {
                SetCurrentUser(employee);
                return true;
            }

            return false;
        }

        private void SetCurrentUser(KmsReportDictionary employee)
        {
            CurrentUser.Region = CmbFilial.Text;
            CurrentUser.FilialCode = CmbFilial.SelectedValue.ToString();
            CurrentUser.Filial = _context.Regions.Single(x => x.Key == CurrentUser.FilialCode).ForeignKey;

            CurrentUser.IdUser = Convert.ToInt32(employee.Key);
            CurrentUser.UserName = CmbLogin.Text;
            CurrentUser.Phone = employee.Value;
            CurrentUser.Email = employee.ForeignKey;
            CurrentUser.Regions = _context.Regions.Where(x => x.Key != "RU").ToList();
            CurrentUser.Users = _context.Users.ToList();
            CurrentUser.ReportTypes = _context.ReportTypes.ToList();               

           var head = _context.Heads.Single(x=>x.FilialCode == CurrentUser.FilialCode);
            CurrentUser.Director = head.Fio;
            CurrentUser.DirectorPhone = head.Phone ?? "";
            CurrentUser.DirectorPosition = head.Position ?? "";
        }

        private List<KmsReportDictionary> GetCurrentUserList(string code) =>
            _context.Users.Where(x => x.ForeignKey == code).ToList();

        private string CheckSavedData()
        {
            var binFormat = new BinaryFormatter();
            string savedLogin = "";
            if (File.Exists(SavedSettings))
            {
                using Stream fStream = File.OpenRead(SavedSettings);
                savedLogin = (string)binFormat.Deserialize(fStream);
            }

            return savedLogin;
        }        
    }
}