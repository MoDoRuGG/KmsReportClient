using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KmsReportClient.Model
{
    public abstract class ElementBase
    {
        public string Name { get; set; }

        public string Index { get; set; }
        public string Description { get; set; }

        public ElementBase()
        {
            Index = "0";
        }

    }
}
