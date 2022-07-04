namespace KmsReportClient.Model.Excel
{
    public class ReportDictionary
    {
        public string TableName { get; set; }
        public int StartRow { get; set; }
        public int EndRow { get; set; }
        public int RowNumIndex { get; set; }
        public int ColumnStartIndex { get; set; }
        public int Index { get; set; }
    }
}
