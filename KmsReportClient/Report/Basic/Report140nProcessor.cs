using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KmsReportClient.DgvHeaderGenerator;
using KmsReportClient.Excel.Creator.Base;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.XML;
using KmsReportClient.Support;
using NLog;

namespace KmsReportClient.Report.Basic
{
    public class Report140nProcessor : AbstractReportProcessor<Report140n>
    {
        StackedHeaderDecorator DgvRender;

        private readonly List<string> headers = new List<string>
{
    // 1. Ведение персонифицированного учета сведений о застрахованных лицах
    "1. Ведение персонифицированного о учета сведений о застрахованных лицах;\r\nКоличество застрахованных лиц СМО согласно региональному сегменту ЕРЗ, содержащих достоверные сведения о застрахованных лицах, за исключением застрахованных лиц, для которых СМО определена в установленном частью 5.1 статьи 16 Федерального закона от 29.11.2010 № 326-ФЗ порядке;\r\nЧЗЛдост\r\n",
    "1. Ведение персонифицированного о учета сведений о застрахованных лицах;\r\nЧисленность застрахованных лиц СМО в субъекте Российской Федерации согласно региональному сегменту ЕРЗ на первое число месяца, следующего за отчетным периодом;\r\nЧЗЛсмо\r\n",
    "1. Ведение персонифицированного о учета сведений о застрахованных лицах;\r\nФормула расчета показателя;\r\nЧЗЛдост*100/ЧЗЛсмо\r\n",

    // 2. Эффективность сопровождения застрахованных лиц при оказании медицинской помощи в рамках базовой программы
    "2. Эффективность сопровождения застрахованных лиц при оказании медицинской помощи в рамках базовой программы;\r\nКоличество письменных обоснованных жалоб и обращений за разъяснениями (консультациями) застрахованных лиц, связанных с вопросами получения медицинской помощи в рамках базовой программы обязательного медицинского страхования, по которым СМО осуществлено сопровождение (предоставлена информация по вопросам ОМС, консультации) за отчетный период;\r\nКСЭрез\r\n",
    "2. Эффективность сопровождения застрахованных лиц при оказании медицинской помощи в рамках базовой программы;\r\nОбщее количество письменных обоснованных жалоб и обращений за разъяснениями (консультациями) застрахованных лиц, связанных с вопросами получения медицинской помощи в рамках базовой программы обязательного медицинского страхования, рассмотренных СМО за отчетный период;\r\nКСЭ\r\n",
    "2. Эффективность сопровождения застрахованных лиц при оказании медицинской помощи в рамках базовой программы;\r\nФормула расчета;\r\nКСЭрез*100/КСЭ\r\n",

    // 3. Эффективность информирования о прохождении диспансеризации взрослого населения
    "3. Эффективность информирования о прохождении диспансеризации взрослого населения;\r\nКоличество застрахованных лиц от 18 лет и старше, проинформированных страховой медицинской организацией о прохождении профилактических мероприятий и прошедших один из видов профилактических мероприятий (диспансеризацию (в том числе углубленную диспансеризацию, для оценки репродуктивного здоровья), профилактические медицинские осмотры) согласно принятым счетам за отчетный период;\r\nППМинфдвн\r\n",
    "3. Эффективность информирования о прохождении диспансеризации взрослого населения;\r\nКоличество застрахованных лиц от 18 лет и старше, индивидуально проинформированных СМО о прохождении профилактических мероприятий за отчетный период;\r\nИидвн\r\n",
    "3. Эффективность информирования о прохождении диспансеризации взрослого населения;\r\nФормула расчета;\r\nППМинфдвн*100/Иидвн\r\n",

    // 4. Эффективность информирования о прохождении диспансерного наблюдения
    "4. Эффективность информирования о прохождении диспансерного наблюдения;\r\nКоличество застрахованных лиц от 18 лет и старше, проинформированных СМО о диспансерном наблюдении и прошедших диспансерное наблюдение за отчетный период;\r\nППМИНФДН\r\n",
    "4. Эффективность информирования о прохождении диспансерного наблюдения;\r\nКоличество застрахованных лиц от 18 лет и старше, индивидуально проинформированных СМО о диспансерном наблюдении за отчетный период;\r\nИидн\r\n",
    "4. Эффективность информирования о прохождении диспансерного наблюдения;\r\nФормула расчета;\r\nППМИНФДН*100/Иидн\r\n",

    // 5. Защита прав застрахованных лиц в судебном и досудебном порядке
    "5. Защита прав застрахованных лиц в судебном и досудебном порядке;\r\nКоличество рассмотренных СМО обоснованных жалоб застрахованных лиц, урегулированных страховой медицинской организацией в досудебном порядке;\r\nКОЖдосуд\r\n",
    "5. Защита прав застрахованных лиц в судебном и досудебном порядке;\r\nКоличество рассмотренных СМО обоснованных жалоб застрахованных лиц, урегулированных СМО в судебном порядке;\r\nКОЖсуд\r\n",
    "5. Защита прав застрахованных лиц в судебном и досудебном порядке;\r\nКоличество поступивших в СМО обоснованных жалоб застрахованных лиц; рассмотренных СМО;\r\nКОЖзл\r\n",
    "5. Защита прав застрахованных лиц в судебном и досудебном порядке;\r\nФормула расчета;\r\n(КОЖдосуд + КОЖсуд)*100/КОЖзл\r\n",

    // 6. Эффективность защиты прав застрахованных лиц
    "6. Эффективность защиты прав застрахованных лиц;\r\nКоличество обоснованных жалоб застрахованных лиц на СМО, поступивших в ТФОМС и СМО от застрахованных лиц напрямую или через иные контрольные органы за отчетный период;\r\nКОЖзлсмо\r\n",
    "6. Эффективность защиты прав застрахованных лиц;\r\nЧисленность застрахованных лиц СМО в субъекте Российской Федерации согласно региональному сегменту ЕРЗ на первое число месяца, следующего за отчетным периодом;\r\nЧЗЛсмо\r\n",
    "6. Эффективность защиты прав застрахованных лиц;\r\nФормула расчета;\r\n100 – КОЖзлсмо*100000/ЧЗЛсмо\r\n",

    // 7. Соблюдение установленного законодательством порядка авансирования медицинских организаций в рамках реализации ТПОМС
    "7. Соблюдение установленного законодательством порядка авансирования медицинских организаций в рамках реализации ТПОМС;\r\nКоличество заявок медицинских организаций на авансирование, направленных в ТФОМС в установленном Правилами ОМС и ПГГ в порядке (без учета заявок медицинских организаций на авансирование с необоснованным превышением установленного Правилами ОМС размера авансирования) за отчетный период;\r\nКЗАсобл\r\n",
    "7. Соблюдение установленного законодательством порядка авансирования медицинских организаций в рамках реализации ТПОМС;\r\nОбщее количество заявок медицинских организаций на авансирование, направленных СМО в ТФОМС за отчетный период;\r\nКЗАвсего\r\n",
    "7. Соблюдение установленного законодательством порядка авансирования медицинских организаций в рамках реализации ТПОМС;\r\nФормула расчета;\r\nКЗАсобл*100/КЗАвсего\r\n",

    // 8. Контроль за использованием медицинскими организациями средств ОМС (дебиторская задолженность МО)
    "8. Контроль за использованием медицинскими организациями средств ОМС (дебиторская задолженность МО);\r\nОбъем средств ОМС, направленных СМО в медицинскую организацию за отчетный период по заявкам на авансирование, не закрытым счетами;\r\nДТ\r\n",
    "8. Контроль за использованием медицинскими организациями средств ОМС (дебиторская задолженность МО);\r\nОбщая сумма средств на оплату медицинской помощи по счетам медицинских организаций, предъявленным к оплате в соответствии с договорами на оказание и оплату медицинской помощи по ОМС за отчетный период;\r\nСчпо\r\n",
    "8. Контроль за использованием медицинскими организациями средств ОМС (дебиторская задолженность МО);\r\nФормула расчета;\r\n1 квартал текущего года: (1-((ДТ/СЧпо) – 1/12*190%))*100,\r\n",

    // 9. Эффективность экспертной деятельности СМО по результатам контроля ТФОМС за качеством проведения СМО контроля объемов, сроков, качества и условий предоставления медицинской помощи
    "9. Эффективность экспертной деятельности СМО по результатам контроля ТФОМС за качеством проведения СМО контроля объемов, сроков, качества и условий предоставления медицинской помощи;\r\nКоличество экспертных заключений СМО по результатам проведенных в соответствии со статьей 40 Федерального закона № 326-ФЗ экспертиз качества медицинской помощи, подтвержденных результатами проведенных ТФОМС повторных экспертиз качества медицинской помощи в рамках осуществляемого ТФОМС контроля за качеством проведения СМО контроля объемов, сроков, качества и условий предоставления медицинской помощи, за отчетный период;\r\nКЭКМПподтв\r\n",
    "9. Эффективность экспертной деятельности СМО по результатам контроля ТФОМС за качеством проведения СМО контроля объемов, сроков, качества и условий предоставления медицинской помощи;\r\nКоличество экспертных заключений по результатам проведенных ТФОМС повторных экспертиз качества медицинской помощи в рамках осуществляемого ТФОМС контроля за качеством проведения СМО контроля объемов, сроков, качества и условий предоставления медицинской помощи за отчетный период;\r\nКЭКМПтфомс\r\n",
    "9. Эффективность экспертной деятельности СМО по результатам контроля ТФОМС за качеством проведения СМО контроля объемов, сроков, качества и условий предоставления медицинской помощи;\r\nФормула расчета;\r\nКЭКМПподтв*100/КЭКМПтфомс\r\n",

    // 10. Качество проводимого СМО контроля объемов, сроков, качества и условий предоставления медицинской помощи (по результатам контроля ТФОМС по претензии МО в рамках обжалования медицинской организацией заключения СМО по результатам контроля)
    "10. Качество проводимого СМО контроля объемов, сроков, качества и условий предоставления медицинской помощи (по результатам контроля ТФОМС по претензии МО в рамках обжалования медицинской организацией заключения СМО по результатам контроля);\r\nКоличество экспертных заключений СМО, подтвержденных результатами проведенной ТФОМС повторной экспертизы качества медицинской помощи по претензии медицинской организации в рамках обжалования заключения СМО по результатам контроля, за отчетный период;\r\nКЗСМОподтв\r\n",
    "10. Качество проводимого СМО контроля объемов, сроков, качества и условий предоставления медицинской помощи (по результатам контроля ТФОМС по претензии МО в рамках обжалования медицинской организацией заключения СМО по результатам контроля);\r\nКоличество экспертных заключений экспертизы качества медицинской помощи, оспариваемых медицинской организацией, поступивших в ТФОМС в рамках обжалования медицинской организацией заключения СМО по результатам контроля, за отчетный период;\r\nКПМОтфомс\r\n",
    "10. Качество проводимого СМО контроля объемов, сроков, качества и условий предоставления медицинской помощи (по результатам контроля ТФОМС по претензии МО в рамках обжалования медицинской организацией заключения СМО по результатам контроля);\r\nФормула расчета;\r\nКЗСМОподтв*100/КПМОтфомс\r\n"
};

        private static readonly Logger Log = LogManager.GetCurrentClassLogger();
        public Report140nProcessor(EndpointSoap inClient, List<KmsReportDictionary> reportsDictionary, DataGridView dgv, ComboBox cmb, TextBox txtb, TabPage page) :
        base(inClient, dgv, cmb, txtb, page,
            XmlFormTemplate.R140n.GetDescription(),
            Log,
            ReportGlobalConst.Report140n,
            reportsDictionary)
        {
            DgvRender = new StackedHeaderDecorator(Dgv);
            InitReport();


        }

        public override AbstractReport CollectReportFromWs(string yymm)
        {
            var request = new GetReportRequest
            {
                Body = new GetReportRequestBody
                {
                    filialCode = FilialCode,
                    yymm = yymm,
                    reportType = ReportType.R140n
                }
            };
            var response = Client.GetReport(request)?.Body?.GetReportResult;
            return response == null ? null : response as Report140n;

        }
        public override void FillDataGridView(string form)
        {
            var report140n = Report.ReportDataList.FirstOrDefault(x => x.Theme == form);
            if (report140n == null || report140n.Data == null)
            {
                return;
            }

            var data = report140n.Data;

            Dgv.Rows[0].Cells[0].Value = data.CZLdost;
            Dgv.Rows[0].Cells[1].Value = data.CZLsmo;
            Dgv.Rows[0].Cells[2].Value = DivideOrZero(data.CZLdost, data.CZLsmo) * 100;

            Dgv.Rows[0].Cells[3].Value = data.KSErez;
            Dgv.Rows[0].Cells[4].Value = data.KSE;
            Dgv.Rows[0].Cells[5].Value = DivideOrZero(data.KSErez, data.KSE) * 100;

            Dgv.Rows[0].Cells[6].Value = data.PPMinadvn;
            Dgv.Rows[0].Cells[7].Value = data.Iidvn;
            Dgv.Rows[0].Cells[8].Value = DivideOrZero(data.PPMinadvn, data.Iidvn) * 100;

            Dgv.Rows[0].Cells[9].Value = data.PPMinfdn;
            Dgv.Rows[0].Cells[10].Value = data.Iidn;
            Dgv.Rows[0].Cells[11].Value = DivideOrZero(data.PPMinfdn, data.Iidn) * 100;

            Dgv.Rows[0].Cells[12].Value = data.KOJdosud;
            Dgv.Rows[0].Cells[13].Value = data.KOJsud;
            Dgv.Rows[0].Cells[14].Value = data.KOJzl;
            Dgv.Rows[0].Cells[15].Value = DivideOrZero((data.KOJdosud + data.KOJsud), data.KOJzl) * 100;

            Dgv.Rows[0].Cells[16].Value = data.KOJzlsmo;
            Dgv.Rows[0].Cells[17].Value = data.CZLsmo;
            Dgv.Rows[0].Cells[18].Value = 100 - DivideOrZero(data.KOJzlsmo * 100000, data.CZLsmo);

            Dgv.Rows[0].Cells[19].Value = data.KZAsobl;
            Dgv.Rows[0].Cells[20].Value = data.KZAvsego;
            Dgv.Rows[0].Cells[21].Value = DivideOrZero(data.KZAsobl, data.KZAvsego) * 100;

            Dgv.Rows[0].Cells[22].Value = data.DT;
            Dgv.Rows[0].Cells[23].Value = data.Scpo;
            var scpoVal = data.Scpo ?? 0;
            Dgv.Rows[0].Cells[24].Value = scpoVal != 0 ? (1 - ((data.DT / scpoVal) - 1M / 12 * 190)) * 100 : 0;

            Dgv.Rows[0].Cells[25].Value = data.KEKMPpodtv;
            Dgv.Rows[0].Cells[26].Value = data.KEKMPtfoms;
            Dgv.Rows[0].Cells[27].Value = DivideOrZero(data.KEKMPpodtv, data.KEKMPtfoms) * 100;

            Dgv.Rows[0].Cells[28].Value = data.KZSMOpodtv;
            Dgv.Rows[0].Cells[29].Value = data.KPMOtfoms;
            Dgv.Rows[0].Cells[30].Value = DivideOrZero(data.KZSMOpodtv, data.KPMOtfoms) * 100;

            SetFormula();
        }

        private decimal DivideOrZero(decimal? numerator, decimal? denominator)
        {
            if (denominator == 0 || denominator == null)
            {
                return 0;
            }
            return (numerator ?? 0) / (denominator ?? 1); // Деление на 1, если denominator null (хотя он уже проверен выше)
        }

        public override void SaveReportDataSourceExcel()
        { }
        public override void SaveReportDataSourceHandle()
        { }

        public void SetFormula()
        {
            try
            {
                // Ячейка 2: (Ячейка 0 / Ячейка 1) * 100
                var cell0 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[0].Value);
                var cell1 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[1].Value);
                Dgv.Rows[0].Cells[2].Value = cell1 != 0 ? Math.Round(cell0 * 100 / cell1, 2) : 0;

                // Ячейка 5: (Ячейка 3 / Ячейка 4) * 100
                var cell3 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[3].Value);
                var cell4 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[4].Value);
                Dgv.Rows[0].Cells[5].Value = cell4 != 0 ? Math.Round(cell3 * 100 / cell4, 2) : 0;

                // Ячейка 8: (Ячейка 6 / Ячейка 7) * 100
                var cell6 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[6].Value);
                var cell7 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[7].Value);
                Dgv.Rows[0].Cells[8].Value = cell7 != 0 ? Math.Round(cell6 * 100 / cell7, 2) : 0;

                // Ячейка 11: (Ячейка 9 / Ячейка 10) * 100
                var cell9 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[9].Value);
                var cell10 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[10].Value);
                Dgv.Rows[0].Cells[11].Value = cell10 != 0 ? Math.Round(cell9 * 100 / cell10, 2) : 0;

                // Ячейка 15: ((Ячейка 12 + Ячейка 13) / Ячейка 14) * 100
                var cell12 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[12].Value);
                var cell13 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[13].Value);
                var cell14 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[14].Value);
                Dgv.Rows[0].Cells[15].Value = cell14 != 0 ? Math.Round((cell12 + cell13) * 100 / cell14, 2) : 0;

                // Ячейка 18: 100 - (Ячейка 16 * 100000 / Ячейка 17)
                var cell16 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[16].Value);
                var cell17 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[17].Value);
                Dgv.Rows[0].Cells[18].Value = cell17 != 0 ? Math.Round(100 - (cell16 * 100000 / cell17), 2) : 100; // Если знаменатель 0, результат 100

                // Ячейка 21: (Ячейка 19 / Ячейка 20) * 100
                var cell19 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[19].Value);
                var cell20 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[20].Value);
                Dgv.Rows[0].Cells[21].Value = cell20 != 0 ? Math.Round(cell19 * 100 / cell20, 2) : 0;

                // Ячейка 24: (1 - ((Ячейка 22 / Ячейка 23) - 1/12 * 190)) * 100
                var cell22 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[22].Value);
                var cell23 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[23].Value);
                Dgv.Rows[0].Cells[24].Value = cell23 != 0 ? Math.Round((1 - ((cell22 / cell23) - 1M / 12 * 190)) * 100, 2) : 0; // 1M для точности деления

                // Ячейка 27: (Ячейка 25 / Ячейка 26) * 100
                var cell25 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[25].Value);
                var cell26 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[26].Value);
                Dgv.Rows[0].Cells[27].Value = cell26 != 0 ? Math.Round(cell25 * 100 / cell26, 2) : 0;

                // Ячейка 30: (Ячейка 28 / Ячейка 29) * 100
                var cell28 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[28].Value);
                var cell29 = GlobalUtils.TryParseDecimal(Dgv.Rows[0].Cells[29].Value);
                Dgv.Rows[0].Cells[30].Value = cell29 != 0 ? Math.Round(cell28 * 100 / cell29, 2) : 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }


        public override void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource)
        {

        }
        public override void InitReport()
        {
            Report = new Report140n { ReportDataList = new Report140nDto[ThemesList.Count], IdType = IdReportType };
            int i = 0;
            foreach (var theme in ThemesList.Select(x => x.Key))
            {
                Report.ReportDataList[i++] = new Report140nDto { Theme = theme };
            }
            SetFormula();
        }
        public override bool IsVisibleBtnDownloadExcel() => false;

        public override bool IsVisibleBtnHandle() => false;

        public override bool IsVisibleBtnSummary() => false;

        public override void MapForAutoFill(AbstractReport report)
        {
            if (report == null)
            {
                return;
            }
            var inReport = report as Report140n;

            var index = Report.ReportDataList.ToList().FindIndex(x => x.Theme == Cmb.Text);
            var inTheme = inReport.ReportDataList.Single(x => x.Theme == Cmb.Text);
            Report.ReportDataList[index] = inTheme;

        }
        public override void SaveToDb()
        {
            SetFormula();
            var request = new SaveReportRequest
            {
                Body = new SaveReportRequestBody
                {
                    filialCode = CurrentUser.FilialCode,
                    idUser = CurrentUser.IdUser,
                    report = Report,
                    yymm = Report.Yymm,
                    reportType = ReportType.R140n
                }
            };
            var response = Client.SaveReport(request).Body.SaveReportResult as Report140n;
            Report.IdFlow = response.IdFlow;
            Report.Status = response.Status;
        }
        public override void ToExcel(string filename, string filialName)
        {

            //var excel = new ExcelCadreCreator(filename, ExcelForm.cadre, Report.Yymm, filialName, Client, FilialCode);
            //excel.CreateReport(Report, null);
        }
        public override string ValidReport() { return ""; }
        protected override void CreateDgvForForm(string form, List<TemplateRow> table)
        {
            Dgv.AllowUserToAddRows = false;
            Dgv.ColumnHeadersVisible = true;

            Dgv.Columns.Clear();
            Dgv.Rows.Clear();

            List<string> columns = null;
            columns = headers;

            foreach (var clmn in columns)
            {
                var column = new DataGridViewTextBoxColumn
                {
                    HeaderText = clmn,
                    DataPropertyName = "Indicator",
                    Name = "Indicator",
                    SortMode = DataGridViewColumnSortMode.NotSortable,
                    DefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.Azure }
                };

                Dgv.Columns.Add(column);
            }

            Dgv.Rows.Add();


            Dgv.Columns[2].ReadOnly =
            Dgv.Columns[5].ReadOnly =
            Dgv.Columns[8].ReadOnly =
            Dgv.Columns[11].ReadOnly =
            Dgv.Columns[15].ReadOnly =
            Dgv.Columns[18].ReadOnly =
            Dgv.Columns[21].ReadOnly =
            Dgv.Columns[24].ReadOnly =
            Dgv.Columns[27].ReadOnly =
            Dgv.Columns[30].ReadOnly =
            true;

            Dgv.AutoSize = true;
            Dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

        }
        protected override void FillReport(string form)
        {
            var report140n = Report.ReportDataList.SingleOrDefault(x => x.Theme == form);
            var row = Dgv.Rows[0];
            report140n.Data = new Report140nDataDto
            {
                CZLdost = GlobalUtils.TryParseDecimal(row.Cells[0].Value),
                CZLsmo = GlobalUtils.TryParseDecimal(row.Cells[1].Value),
                KSErez = GlobalUtils.TryParseDecimal(row.Cells[3].Value),
                KSE = GlobalUtils.TryParseDecimal(row.Cells[4].Value),
                PPMinadvn = GlobalUtils.TryParseDecimal(row.Cells[6].Value),
                Iidvn = GlobalUtils.TryParseDecimal(row.Cells[7].Value),
                PPMinfdn = GlobalUtils.TryParseDecimal(row.Cells[9].Value),
                Iidn = GlobalUtils.TryParseDecimal(row.Cells[10].Value),
                KOJdosud = GlobalUtils.TryParseDecimal(row.Cells[12].Value),
                KOJsud = GlobalUtils.TryParseDecimal(row.Cells[13].Value),
                KOJzl = GlobalUtils.TryParseDecimal(row.Cells[14].Value),
                KOJzlsmo = GlobalUtils.TryParseDecimal(row.Cells[16].Value),
                KZAsobl = GlobalUtils.TryParseDecimal(row.Cells[19].Value),
                KZAvsego = GlobalUtils.TryParseDecimal(row.Cells[20].Value),
                DT = GlobalUtils.TryParseDecimal(row.Cells[22].Value),
                Scpo = GlobalUtils.TryParseDecimal(row.Cells[23].Value),
                KEKMPpodtv = GlobalUtils.TryParseDecimal(row.Cells[25].Value),
                KEKMPtfoms = GlobalUtils.TryParseDecimal(row.Cells[26].Value),
                KZSMOpodtv = GlobalUtils.TryParseDecimal(row.Cells[28].Value),
                KPMOtfoms = GlobalUtils.TryParseDecimal(row.Cells[29].Value)
            };
        }
    }
}
