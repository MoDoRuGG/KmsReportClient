﻿using System.Collections.Generic;
using KmsReportClient.External;
using KmsReportClient.Model.Enums;
using Microsoft.Office.Interop.Excel;

namespace KmsReportClient.Excel.Creator.Consolidate
{
    class ExcelFFOMSMonthlyVolCreator : ExcelBaseCreator<FFOMSMonthlyVol>
    {
        public ExcelFFOMSMonthlyVolCreator(
            string filename,
            string header,
            string filialName) : base(filename, ExcelForm.MonthlyVol, header, filialName, true) { }

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
            {"RU-MOS","Дирекция по работе в Московской области ООО Капитал МС"},
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

        protected override void FillReport(FFOMSMonthlyVol report, FFOMSMonthlyVol yearReport)
        {
            ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];
            if (region_name.TryGetValue(report.Filial, out string regionName))
            {
                ObjWorkSheet.Cells[1, 1] = regionName;
            }
            FillSKP(report);
            FillSDP(report);
            FillAPP(report);
            FillSMP(report);
        }

        private void FillSKP(FFOMSMonthlyVol report)
        {
            int currentIndex = 9;
            foreach (var data in report.FFOMSMonthlyVol_SKP)
            {
                ObjWorkSheet.Cells[currentIndex, 2] = data.CountSluch;
                ObjWorkSheet.Cells[currentIndex, 3] = data.CountAppliedSluch;
                ObjWorkSheet.Cells[currentIndex, 6] = data.CountSluchMEE;
                ObjWorkSheet.Cells[currentIndex++, 10] = data.CountSluchEKMP;
            }
        }

        private void FillSDP(FFOMSMonthlyVol report)
        {
            int currentIndex = 23;
            foreach (var data in report.FFOMSMonthlyVol_SDP)
            {
                ObjWorkSheet.Cells[currentIndex, 2] = data.CountSluch;
                ObjWorkSheet.Cells[currentIndex, 3] = data.CountAppliedSluch;
                ObjWorkSheet.Cells[currentIndex, 6] = data.CountSluchMEE;
                ObjWorkSheet.Cells[currentIndex++, 10] = data.CountSluchEKMP;
            }
        }

        private void FillAPP(FFOMSMonthlyVol report)
        {
            int currentIndex = 37;
            foreach (var data in report.FFOMSMonthlyVol_APP)
            {
                ObjWorkSheet.Cells[currentIndex, 2] = data.CountSluch;
                ObjWorkSheet.Cells[currentIndex, 3] = data.CountAppliedSluch;
                ObjWorkSheet.Cells[currentIndex, 6] = data.CountSluchMEE;
                ObjWorkSheet.Cells[currentIndex++, 10] = data.CountSluchEKMP;
            }
        }

        private void FillSMP(FFOMSMonthlyVol report)
        {
            int currentIndex = 51;
            foreach (var data in report.FFOMSMonthlyVol_SMP)
            {
                ObjWorkSheet.Cells[currentIndex, 2] = data.CountSluch;
                ObjWorkSheet.Cells[currentIndex, 3] = data.CountAppliedSluch;
                ObjWorkSheet.Cells[currentIndex, 6] = data.CountSluchMEE;
                ObjWorkSheet.Cells[currentIndex++, 10] = data.CountSluchEKMP;
            }
        }
    }
}
