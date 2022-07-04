using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KmsReportClient.Model.Enums
{
    public enum TableElement
    {
        [Description("Столбец")]
        Column,
        [Description("Строка")]
        Row,
        [Description("Вкладка")]
        Page,
        [Description("Группа")]
        Group
    }
}
