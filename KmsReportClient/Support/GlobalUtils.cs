using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using KmsReportClient.External;
using KmsReportClient.Global;

namespace KmsReportClient.Support
{
    static class GlobalUtils
    {


        public static int TryParseInt(object str) =>
            int.TryParse(str != null ? str.ToString() : "", out var number)
                ? number
                : 0;

        public static decimal TryParseDecimal(object str) =>
            decimal.TryParse(
                str != null ? str.ToString().Replace(".", ",") : "",
                out var number)
                ? number
                : 0;

        public static void AddValueInTextBox(ComboBox cmb, TextBox txtb, bool lookValue, bool all)
        {
            if (!lookValue && cmb.SelectedValue == null)
            {
                return;
            }

            var value = lookValue ? cmb.Text : cmb.SelectedValue.ToString();
            var parts = txtb.Text.Replace(", ", ",").Split(',').ToList();

            if (all)
            {
                txtb.Clear();
                foreach (var region in cmb.Items)
                {
                    value = (region as KmsReportDictionary).Value;
                    if (value != "Все филиалы")
                        parts.Add(value);
                }

            }
            else
            {
                parts.Add(value);

            }
            if (parts[0].Length == 0)
            {
                parts.RemoveAt(0);
            }


            parts = parts.Distinct().ToList();

            txtb.Text = string.Join(", ", parts).Replace(", ", ",");

        }

        public static void DeleteValueFromTextBox(TextBox textBox)
        {
            var parts = textBox.Text.Replace(", ", ",").Split(',').ToList();
            if (parts.Count > 0)
            {
                parts.RemoveAt(parts.Count - 1);
                textBox.Text = string.Join(", ", parts);
            }
        }

        public static string GetSerializeName(string serializeName, string yymm) =>
            "Serializable\\" + CurrentUser.FilialCode + serializeName + yymm +
            ".dat";

        public static void OpenFileOrDirectory(string filename)
        {
            var dialogResult = MessageBox.Show("Показать результаты?", "Информация", MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);

            if (dialogResult == DialogResult.Yes)
            {
                Process.Start(filename);
            }
        }
    }
}