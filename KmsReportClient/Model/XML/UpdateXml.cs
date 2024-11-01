using System.Collections.Generic;
using System.Xml.Serialization;

namespace KmsReportClient.Model.XML
{
    [XmlRoot("KMS_REPORT")]
    public class UpdateXml
    {
        [XmlElement("VERSION")] public string Version { get; set; }
        [XmlElement("FILE")] public List<RemoteFile> RemoteFiles { get; set; }
    }

    [XmlRoot("FILE")]
    public class RemoteFile
    {
        [XmlElement("NAME")]
        public string Name { get; set; }
        [XmlElement("IS_NEED_DOWNLOAD")]
        public bool IsNeedDownload { get; set; }
        [XmlElement("HASH")]
        public string Hash { get; set; }
    }
}