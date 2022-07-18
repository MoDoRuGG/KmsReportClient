using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Xml.Serialization;
using KmsReportClient.External;
using KmsReportClient.Global;
using KmsReportClient.Model.Enums;
using KmsReportClient.Model.XML;
using KmsReportClient.Support;
using NLog;

namespace KmsReportClient.Service
{
    internal class ApplicationUpdater
    {
        private const string ApplicationName = "KmsReportClient.exe";
        private const string TempApplicationName = "Temp//KmsReportClient.exe";
        private const string TemplatesFolder = "Template\\";

        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        private readonly EndpointSoapClient _client;
        private readonly FileProcessor _ftpProcessor;

        public ApplicationUpdater(FileProcessor ftpProcessor, EndpointSoapClient client)
        {
            this._ftpProcessor = ftpProcessor;
            this._client = client;
        }


        public void GetDll()
        {
            //скачивание dll
            string[] dllFiles = _ftpProcessor.GetDllFileNames(_client);
            if (dllFiles.Length != 0)
            {
                foreach (var fileName in dllFiles)
                {

                    _ftpProcessor.DownloadDll(_client, fileName);

                }
            }
        }

        public void UpdateApp(bool idApplicationStart)
        {
            var currentVersion = Convert.ToDouble(Application.ProductVersion.Replace(".", ""));
            try
            {
                var updateFile = _ftpProcessor.DownloadFileFromWs(XmlFormTemplate.UpdateXml.GetDescription(), "", "", _client);
                var updateXml = ReadUpdateXml(updateFile);
                if (updateXml == null)
                {
                    var errorMessage = "Файл с информацией об обновлении некорректен.";
                    Log.Error(errorMessage);
                    MessageBox.Show(errorMessage, @"Ошибка обновления",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                File.Delete("Temp\\upd_ver.xml");

                foreach (var file in updateXml.RemoteFiles.Where(file => !File.Exists(GlobalConst.TempFolder + file.Name)))
                {
                    DownloadExcelTemplate(file.Name);
                }
                var versionRemoteT = updateXml.Version;
                var versionRemote = Convert.ToDouble(updateXml.Version.Replace(".", ""));

                if (versionRemote > currentVersion)
                {
                    foreach (var file in updateXml.RemoteFiles.Where(f => f.IsNeedDownload))
                    {
                        if (file.Name == XmlFormTemplate.TextMail.GetDescription() && File.Exists(TemplatesFolder + file.Name))
                        {
                            continue;
                        }
                        DownloadExcelTemplate(file.Name);
                    }

                    var message = $"Обновление до версии {versionRemote}";
                    Log.Info(message);
                    MessageBox.Show(@$"Приложение будет автоматически обновлено до версии {versionRemoteT} и перезапущено!", @"Внимание",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);


                    ////скачивание dll
                    //string[] dllFiles = _ftpProcessor.GetDllFileNames(_client);
                    //foreach (var fileName in dllFiles)
                    //{
                    //    _ftpProcessor.DownloadDll(_client, fileName);

                    //};
                  

                    _ftpProcessor.DownloadFileFromWs(ApplicationName, "", "", _client);
                    Process.Start("Updater.exe", $"{TempApplicationName} {ApplicationName}");
                    Process.GetCurrentProcess().Kill();
                }
                else
                {
                    Log.Info($"Обновление не требуется. Актуальная версия приложения: {currentVersion}");
                    if (!idApplicationStart)
                    {
                        var message = $"Текущая версия приложения: {Application.ProductVersion}. Обновлений нет!";
                        MessageBox.Show(message, @"Информация об обновлении", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                var message = $"Ошибка обновления. {ex}";
                Log.Error(ex, "Ошибка обновления");
                MessageBox.Show(message, @"Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DownloadExcelTemplate(string fileName)
        {
            var downloadingFilePath = _ftpProcessor.DownloadFileFromWs(fileName, "Template", "", _client);

            if (downloadingFilePath != null)
            {
                File.Delete(TemplatesFolder + fileName);
                File.Move(downloadingFilePath, TemplatesFolder + fileName);
            }
        }

        private UpdateXml ReadUpdateXml(string xmlPath)
        {
            var xmlDoc = XDocument.Load(xmlPath);
            var xmlSerializer = new XmlSerializer(typeof(UpdateXml));
            if (xmlDoc.Root != null)
            {
                using var reader = xmlDoc.Root.CreateReader();
                return (UpdateXml)xmlSerializer.Deserialize(reader);
            }

            return null;
        }
    }
}