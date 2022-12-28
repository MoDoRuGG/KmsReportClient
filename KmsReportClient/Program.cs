using System;
using System.Windows.Forms;
using KmsReportClient.Forms;

namespace KmsReportClient
{
    static class Program
    {
        /// <summary>
        ///     Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new AuthorizationForm());

            if (AuthorizationForm.Status)
            {
                Application.Run(new MainForm());
            }
        }
    }
}