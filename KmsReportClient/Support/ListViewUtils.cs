using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KmsReportClient.Support
{
    public static class ListViewUtils
    {
        public static void AddItem(ListView listView, string itemName)
        {
            int number = GetNumberValueLastItem(listView) + 1;
            string[] itemData = { number.ToString(), itemName };
            var listViewItem = new ListViewItem(itemData);
            listView.Items.Add(listViewItem);

        }
        public static int GetNumberValueLastItem(ListView listView)
        {
            int listViewCount = listView.Items.Count;

            if (listViewCount == 0)
                return 0;

            return Convert.ToInt32(listView.Items[listViewCount - 1].SubItems[0].Text);
        }
    }
}
