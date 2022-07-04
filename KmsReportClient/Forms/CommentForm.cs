using System;
using System.Linq;
using System.Windows.Forms;
using KmsReportClient.External;
using KmsReportClient.Global;

namespace KmsReportClient.Forms
{
    public partial class CommentForm : Form
    {
        private readonly EndpointSoap _client;
        private readonly int _idReport;

        public CommentForm(EndpointSoap inClient, AbstractReport report)
        {
            InitializeComponent();

            _client = inClient;
            _idReport = report.IdFlow;
        }

        private void BtnDo_Click(object sender, EventArgs e)
        {
            var com = TxtbComment.Text;
            if (!string.IsNullOrEmpty(com))
            {
                var request = new AddCommentRequest
                {
                    Body = new AddCommentRequestBody
                    {
                        comment = com,
                        idReport = _idReport,
                        idEmp = CurrentUser.IdUser
                    }
                };
                _client.AddComment(request);

                GetComment();
                TxtbComment.Clear();
            }
            else
            {
                MessageBox.Show("Для отправки сообщения необходимо его заполнить", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void GetComment()
        {
            var request = new GetCommentsRequest
            {
                Body = new GetCommentsRequestBody { idReport = _idReport, filialCode = CurrentUser.FilialCode }
            };
            var comments = _client.GetComments(request);
            var commentsList = comments.Body.GetCommentsResult;
            if (commentsList.Length > 0)
            {
                commentsList = commentsList.OrderBy(x => x.DateIns).ToArray();
                string text = "";
                foreach (var com in commentsList)
                {
                    text += $"{com.DateIns.ToShortDateString()} {com.Name}: {com.Comment}" + Environment.NewLine;
                }

                TxtbChat.Text = text;
            }
        }

        private void BtnRefresh_Click(object sender, EventArgs e) =>
            GetComment();

        private void CommentForm_Load(object sender, EventArgs e) =>
            GetComment();

        private void BtnClear_Click(object sender, EventArgs e) =>
            TxtbComment.Clear();

        private void BtnClose_Click(object sender, EventArgs e) =>
            Close();
    }
}