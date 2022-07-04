using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Xml.Serialization;
using KmsReportClient.Model.XML;
using KmsReportClient.Utils;

namespace KmsReportClient.Report
{
    [Serializable]
    public abstract class AbstractReportProcessor
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string SmallName { get; set; }
        public string SerializeName { get; set; }
        public string OldTheme { get; set; }
        public string Yymm { get; set; }
        public string FilialCode { get; set; }
        public string FilialName { get; set; }
        public string TemplateFormName { get; set; }

        public DateTime Created { get; set; }
        public int IdEmployee { get; set; }
        public DateTime? Updated { get; set; }
        public int IdEmployeeUpd { get; set; }

        public DateTime? DateToCo { get; set; }
        public int UserToCo { get; set; }
        public DateTime? RefuseDate { get; set; }
        public int RefuseUser { get; set; }
        public DateTime? DateIsDone { get; set; }
        public int UserSubmit { get; set; }

        public int Version { get; set; }
        public string Scan { get; set; }

        public bool Refuse { get; set; }
        public bool Submit { get; set; }
        public bool IsDone { get; set; }

        [NonSerialized]
        protected TemplateForm templateForm;
        [NonSerialized]
        protected External.EndpointSoap client;
        [NonSerialized]
        private static readonly NLog.Logger Log = NLog.LogManager.GetCurrentClassLogger();

        public abstract void FillReport(DataGridView dgvReport, string form);
        public abstract void FillDataGridView(DataGridView dgvReport, string form);
        public abstract void SaveToDb();
        public abstract void FindReports(List<string> filialList, string yymmStart, string yymmEnd);
        public abstract void ToExcel(string filename);
        public abstract string ValidReport();
        protected abstract void CreateDgvForForm(DataGridView dgvReport, string form, List<TemplateRow> table);

        public void MapFromExternal(External.AbstractReport inReport) =>
            MappingUtils.MapFromExternal(this, inReport);

        public void SetClient(External.EndpointSoap inClient) =>
            client = inClient;

        public void Serialize(string yymm)
        {
            var filename = GlobalUtils.GetSerializeName(SerializeName, yymm);
            var binFormat = new BinaryFormatter();
            using (Stream fStream = new FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                binFormat.Serialize(fStream, this);
            }
        }

        public void CreateReportForm(DataGridView dgvReport, string form)
        {
            dgvReport.AutoGenerateColumns = false;
            dgvReport.AllowUserToAddRows = false;
            dgvReport.AutoSize = false;
            dgvReport.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dgvReport.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dgvReport.Columns.Clear();
            dgvReport.Rows.Clear();

            if (templateForm == null)
            {
                templateForm = ReadTemplateXml();
            }
            var table = templateForm.tables
                    .Where(x => x.Name == form)
                    .SelectMany(x => x.Rows)
                    .ToList();

            CreateDgvForForm(dgvReport, form, table);
        }

        public List<KmsReportDictionary> GetForms() =>
            templateForm.tables
                .Select(x => new KmsReportDictionary(x.Name, x.TableDescription))
                .ToList();

        public string GetReportInfo()
        {
            string info = $"{SmallName}; Период: {Yymm}; " + Environment.NewLine;

            if (IdEmployee != 0)
            {
                info += $"Дата создания: {Created.ToShortDateString()} ";
                info += $"Пользователь: {GetUser(IdEmployee)}; " + Environment.NewLine;
            }

            if (IdEmployeeUpd != 0 && Updated != null)
            {
                info += $"Дата обновления: {Updated.Value.ToShortDateString()} ";
                info += $"Пользователь: {GetUser(IdEmployeeUpd)}; " + Environment.NewLine;
            }

            info += "Наличие скана: " + (!string.IsNullOrEmpty(Scan) ? "Да; " : "Нет; ") + Environment.NewLine;
            if (UserToCo != 0 && DateToCo != null)
            {
                info += $"Дата направления в ЦО: {DateToCo.Value.ToShortDateString()} ";
                info += $"Пользователь: {GetUser(UserToCo)}; " + Environment.NewLine;
            }

            if (RefuseUser != 0 && RefuseDate != null)
            {
                info += $"Дата возврата отчета на доработку: {RefuseDate.Value.ToShortDateString()} ";
                info += $"Пользователь: {GetUser(RefuseUser)}; " + Environment.NewLine;
            }

            if (UserSubmit != 0 && DateIsDone != null)
            {
                info += $"Дата утверждения: {DateIsDone.Value.ToShortDateString()} ";
                info += $"Пользователь: {GetUser(UserSubmit)}; " + Environment.NewLine;
            }
            return info;
        }

        public bool SaveScan(string inUri)
        {
            try
            {
                var request = new External.SaveScanRequest
                {
                    Body = new External.SaveScanRequestBody
                    {
                        idUser = CurrentUser.IdUser,
                        filialCode = CurrentUser.FilialCode,
                        report = MappingUtils.MapToExternalAbstractReport(this),
                        uri = inUri
                    }
                };
                var response = client.SaveScan(request);
                return response.Body.SaveScanResult;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка сохранения скана в БД");
                throw;
            }
        }

        public bool SubmitReport()
        {
            try
            {
                var request = new External.SubmitReportRequest
                {
                    Body = new External.SubmitReportRequestBody
                    {
                        idUser = CurrentUser.IdUser,
                        filialCode = CurrentUser.FilialCode,
                        report = MappingUtils.MapToExternalAbstractReport(this)
                    }
                };
                var response = client.SubmitReport(request);
                return response.Body.SubmitReportResult;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка утверждения филиалом отчета");
                throw;
            }
        }

        public bool RefuseReport(string filialCode)
        {
            try
            {
                var request = new External.RefuseReportRequest
                {
                    Body = new External.RefuseReportRequestBody
                    {
                        idUser = CurrentUser.IdUser,
                        filialCode = filialCode,
                        report = MappingUtils.MapToExternalAbstractReport(this)
                    }
                };
                var response = client.RefuseReport(request);
                return response.Body.RefuseReportResult;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка возврата отчета на доработку");
                throw;
            }
        }

        public bool DoneReport(string filialCode)
        {
            try
            {
                var request = new External.DoneReportRequest
                {
                    Body = new External.DoneReportRequestBody
                    {
                        idUser = CurrentUser.IdUser,
                        filialCode = filialCode,
                        report = MappingUtils.MapToExternalAbstractReport(this)
                    }
                };
                var response = client.DoneReport(request);
                return response.Body.DoneReportResult;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка утверждения отчета");
                throw;
            }
        }

        protected bool IsNotNeedFillRow(string form, string rowNum)
        {
            return templateForm.tables
                .Where(x => x.Name == form)
                .SelectMany(x => x.Rows)
                .Single(x => x.Num == rowNum)
                .Exclusion;
        }

        protected TemplateForm ReadTemplateXml()
        {
            try
            {
                var xmlDoc = XDocument.Load(TemplateFormName);
                var xmlSerializer = new XmlSerializer(typeof(TemplateForm));
                if (xmlDoc.Root != null)
                {
                    using (var reader = xmlDoc.Root.CreateReader())
                    {
                        return (TemplateForm)xmlSerializer.Deserialize(reader);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error($"Ошибка получения данных из шаблона для формы: {ex}");
            }
            return null;
        }

        private string GetUser(int key)
        {
            try
            {
                var request = new External.GetUserByCodeRequest
                {
                    Body = new External.GetUserByCodeRequestBody
                    {
                        key = key
                    }
                };
                return client.GetUserByCode(request).Body.GetUserByCodeResult;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Ошибка получения пользователя");
                throw;
            }
        }
    }
}
