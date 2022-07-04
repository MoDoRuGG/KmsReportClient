using System;
using System.IO;
using KmsReportClient.External;
using NLog;

namespace KmsReportClient.Service
{
    class FileProcessor
    {
        private static readonly Logger Log = LogManager.GetCurrentClassLogger();

        public void UploadFileToWs(string filename, string filial, string finalName,
            EndpointSoapClient client)
        {
            Log.Info($"Start uploading file to ftp {filename}");

            var fInfo = new FileInfo(filename);
            long numBytes = fInfo.Length;
            double dLen = (double)fInfo.Length / (1024 * 1024);
            if (dLen > 20)
            {
                throw new Exception("Максимальный размер файла не должен превышать 20МБ");
            }

            try
            {
                byte[] data;
                using var fStream = new FileStream(filename, FileMode.Open, FileAccess.Read);
                using var br = new BinaryReader(fStream);
                data = br.ReadBytes((int)numBytes);

                client.UploadFile(data, finalName, filial);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error uploading file to ftp");
                throw;
            }
        }

        public string DownloadFileFromWs(string scan, string folder, string filial, EndpointSoapClient client)
        {
            var parts = scan.Split('/');
            var filename = parts.Length > 1 ? parts[parts.Length - 1] : parts[0];

            Log.Info($"Start downloading file from ftp {filename}");

            var inputFilePath = string.IsNullOrEmpty(folder) ? filename : $"{folder}\\{filename}";
            var outputFileName = $"Temp\\{filename}";
            try
            {
                var data = client.DownloadFile(inputFilePath, filial);
                using (var fs = new FileStream(outputFileName, FileMode.Create))
                {
                    var ms = new MemoryStream(data);
                    ms.WriteTo(fs);
                    ms.Close();
                }

                return outputFileName;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error downloading file from ftp");
                throw;
            }
        }

        public void DownloadDll(EndpointSoapClient client,string fileName)
        {

            Log.Info($"Start downloading file from ftp {fileName}");        
           
            try
            {
                var data = client.DownloadDllFile(fileName);
                using (var fs = new FileStream(System.IO.Directory.GetCurrentDirectory() + "/" + Path.GetFileName(fileName) , FileMode.Create))
                {
                    var ms = new MemoryStream(data);
                    ms.WriteTo(fs);
                    ms.Close();
                }
            
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error downloading file from ftp");
                throw;
            }
        }

        public string[] GetDllFileNames(EndpointSoapClient client)
        {
            ArrayOfString filesNames = new ArrayOfString();
            Log.Info($"Старт получения DLL");
            try
            {
                 filesNames = client.GetDllFileNames();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error get file names from FTP server");
                throw;
            }

            return filesNames.ToArray();
        }
    }
}