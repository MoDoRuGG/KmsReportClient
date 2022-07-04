using System.Collections.Generic;
using System.Drawing;
using KmsReportClient.External;

namespace KmsReportClient.Global
{
    public static class GlobalConst
    {
        public static readonly string[] Months = { "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль",
            "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
        };

        public static readonly string[] Periods = { "1й квартал", "1е полугодие", "9 месяцев", "Год" };
        public static readonly string[] PeriodsQ = { "1 квартал", "2 квартал", "3 квартал", "4 квартал" };

        public static readonly Dictionary<string, string> MonthsWithNumber = new Dictionary<string, string>
        {
             {"01","январь" },
             {"02","февраль" },
             {"03","март" },
             {"04","апрель" },
             {"05","май" },
             {"06","июнь" },
             {"07","июль" },
             {"08","август" },
             {"09","сентябрь" },
             {"10","октябрь" },
             {"11","ноябрь" },
             {"12","декабрь" },
        };




        public const string TempFolder = @"Template\";

        public static Color ColorIsDone = Color.YellowGreen;
        public static Color ColorRefuse = Color.Tomato;
        public static Color ColorSubmit = Color.PaleGreen;
        public static Color ColorScan = Color.FromArgb(255, 255, 128);
        public static Color ColorBd = Color.LightYellow;

        public static List<KmsReportDictionary> EmailList = new List<KmsReportDictionary> {

            new KmsReportDictionary {Key = "management_of_economics@kapmed.ru",Value = "Директора" },
            new KmsReportDictionary {Key = "managers.zpz.ekmp@kapmed.ru",Value = "Начальники отделов ЗПЗ и ЭКМП филиалов" },
            new KmsReportDictionary {Key = "managers.it@kapmed.ru",Value = "Начальники отделов ИТ филиалов" },
            new KmsReportDictionary {Key = "management_of_economics@kapmed.ru",Value = "Начальники экономических отделов филиалов" },           
            new KmsReportDictionary {Key = "management_of_economics@kapmed.ru",Value = "Бухгалтерия"},

        };

        public static List<KmsReportDictionary> FilterList = new List<KmsReportDictionary> {
            new KmsReportDictionary {Key = "Saved", Value = "Сохранен в БД"},
            new KmsReportDictionary {Key = "Scan", Value = "Направлен скан"},
            new KmsReportDictionary {Key = "Submit", Value = "Направлен в ЦО"},
            new KmsReportDictionary {Key = "Refuse", Value = "На доработке"},
            new KmsReportDictionary {Key = "Done", Value = "Принят в ЦО"}
           
        };

        public static ReportStatus[] SuccessStatuses = { ReportStatus.Done, ReportStatus.Submit };
    }
}