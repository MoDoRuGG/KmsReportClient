using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KmsReportClient.External;
using KmsReportClient.Model;
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.Excel;
using KmsReportClient.Support;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Base
{
    public class ExcelMonthlyVolCreator : ExcelBaseCreator<ReportMonthlyVol>
    {

        private readonly List<ReportDictionary> _MonVolDictionaries = new List<ReportDictionary> {

            new ReportDictionary {TableName = "Стационарная помощь", StartRow = 9, EndRow = 21, Index = 1},
            new ReportDictionary {TableName = "Дневной стационар", StartRow = 23, EndRow = 35, Index = 2},
            new ReportDictionary {TableName = "АПП", StartRow = 37, EndRow = 49, Index = 3},
            new ReportDictionary {TableName = "Скорая медицинская помощь", StartRow = 51, EndRow = 63, Index = 4},
        };

        private static readonly Dictionary<string, string> region_name = new Dictionary<string, string>
        {
            {"RU-AL","АСП ООО «Капитал МС» - Филиал в Республике Алтай"},
            {"RU-ALT","АСП ООО «Капитал МС» - Филиал в Алтайском крае"},
            {"RU-ARK","АСП ООО «Капитал МС» - Филиал в Архангельской области "},
            {"RU-BA","АСП ООО «Капитал МС» - Филиал в Республике Башкортостан"},
            {"RU-BU","АСП ООО «Капитал МС» - Филиал в Республике Бурятия"},
            {"RU-KB","АСП ООО «Капитал МС» - Филиал в Кабардино-Балкарской Республике"},
            {"RU-KDA","АСП ООО «Капитал МС» - Филиал в Краснодарском крае"},
            {"RU-KGD","АСП ООО «Капитал МС» - Филиал в Калининградской области"},
            {"RU-KGN","АСП ООО «Капитал МС» - Филиал в Курганской области"},
            {"RU-KHA","АСП ООО «Капитал МС» - Филиал в Хабаровском крае"},
            {"RU-KHM","АСП ООО «Капитал МС» - Филиал в Ханты-Мансийском Автономном округе-Югре"},
            {"RU-KIR","АСП ООО «Капитал МС» - Филиал в Кировской области"},
            {"RU-KO","АСП ООО «Капитал МС» - Филиал в Республике Коми"},
            {"RU-KOS","АСП ООО «Капитал МС» - Филиал в Костромской области"},
            {"RU-LEN","АСП ООО «Капитал МС» - Филиал в г.Санкт-Петербурге и Ленинградской области"},
            {"RU-LIP","АСП ООО «Капитал МС» - Филиал в Липецкой области"},
            {"RU-MO","АСП ООО «Капитал МС» - Филиал в Республике Мордовия"},
            {"RU-MOS","Дирекция по работе в Московской области ООО «Капитал МС»"},
            {"RU-MOW","АСП ООО «Капитал МС» - Филиал в г. Москве"},
            {"RU-NEN","АСП ООО «Капитал МС» - Филиал в Ненецком Автономном Округе"},
            {"RU-NIZ","АСП ООО «Капитал МС» - Филиал в Нижегородской области"},
            {"RU-OMS","АСП ООО «Капитал МС» - Филиал в Омской области"},
            {"RU-ORE","АСП ООО «Капитал МС» - Филиал в Оренбургской области"},
            {"RU-PER","АСП ООО «Капитал МС» - Филиал в Пермском крае"},
            {"RU-PNZ","АСП ООО «Капитал МС» - Филиал в Пензенской области"},
            {"RU-ROS","АСП ООО «Капитал МС» - Филиал в Ростовской области"},
            {"RU-RYA","АСП ООО «Капитал МС» - Филиал в Рязанской области"},
            {"RU-SA","АСП ООО «Капитал МС» - Филиал в Республике Саха (Якутия)"},
            {"RU-SAR","АСП ООО «Капитал МС» - Филиал в Саратовской области"},
            {"RU-SE","АСП ООО «Капитал МС» - Филиал в Республике Северная Осетия-Алания"},
            {"RU-SMO","АСП ООО «Капитал МС» - Филиал в Смоленской области"},
            {"RU-SPE","АСП ООО «Капитал МС» - Филиал в г.Санкт-Петербурге и Ленинградской области"},
            {"RU-TUL","АСП ООО «Капитал МС» - Филиал в Тульской области"},
            {"RU-TVE","АСП ООО «Капитал МС» - Филиал в Тверской области"},
            {"RU-TY","АСП ООО «Капитал МС» - Филиал в Республике Тыва"},
            {"RU-TYU","АСП ООО «Капитал МС» - Филиал в Тюменской области"},
            {"RU-UD","АСП ООО «Капитал МС» - Филиал в Удмуртской Республике"},
            {"RU-ULY","АСП ООО «Капитал МС» - Филиал в Ульяновской области"},
            {"RU-VGG","АСП ООО «Капитал МС» - Филиал в Волгоградской области"},
            {"RU-VLA","АСП ООО «Капитал МС» - Филиал во Владимирской области"},
            {"RU-YAR","АСП ООО «Капитал МС» - Филиал в Ярославской области"},
            {"RU-YEV","АСП ООО «Капитал МС» - Филиал в Еврейской Автономной Области"},
            {"RU", "Центральный офис ООО «Капитал МС»"}
        };


        public ExcelMonthlyVolCreator(
          string filename,
          ExcelForm reportName,
          string header,
          string filialName) : base(filename, reportName, header, filialName, false)
        {
        }

        protected override void FillReport(ReportMonthlyVol report, ReportMonthlyVol yearReport)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            if (region_name.TryGetValue(FilialName, out string regionName))
            {
                ObjWorkSheet.Cells[1, 1] = regionName;
            }
            foreach (var themeData in report.ReportDataList.OrderBy(x => x.Theme))
            {
                var dict = _MonVolDictionaries.Single(x => x.TableName == themeData.Theme);
                var data = themeData.Data;
                FillTable(data, dict.StartRow, dict.EndRow, themeData.Theme);
            }
        }




        private void FillTable(ReportMonthlyVolDataDto[] data, int startRowIndex, int endRowIndex, string form)
        {
            int j = startRowIndex;
            if (data != null)
            {
                foreach (var row in data)
                {
                    ObjWorkSheet.Cells[j, 2] = row.CountSluch;
                    ObjWorkSheet.Cells[j, 3] = row.CountAppliedSluch;
                    ObjWorkSheet.Cells[j, 6] = row.CountSluchMEE;
                    ObjWorkSheet.Cells[j++, 10] = row.CountSluchEKMP;
                    if (j == endRowIndex)
                    break;
                }
            };
        }
    }
}
