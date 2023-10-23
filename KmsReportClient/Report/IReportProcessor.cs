using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using KmsReportClient.External;

namespace KmsReportClient.Report
{
   public interface IReportProcessor
    {
        AbstractReport Report { get; set; }
        Color ColorReport { get; set; }
        bool HasReport { get; set; }
        List<KmsReportDictionary> ThemesList { get; }
        TabPage Page { get; }
        string SmallName { get; }
        string FilialName { get; }
        string FilialCode { get; set; }
        string OldTheme { get; set; }
        void CreateTotalColumn();
        void SetTotalColumn();
        void MapReportFromDgv(string form);
        void FillDataGridView(string form);
        void SaveToDb();
        void SaveReportDataSourceHandle();
        void SaveReportDataSourceExcel();
        void FindReports(List<string> filialList, string yymmStart, string yymmEnd, ReportStatus status, DataSource datasource);
        void ToExcel(string filename, string filialName);
        string ValidReport();
        bool IsVisibleBtnDownloadExcel();
        bool IsVisibleBtnHandle();
        bool IsVisibleBtnSummary();

        void InitReport();
        void SetReadonlyForDgv(bool isReadonly);
        AbstractReport CollectReportFromWs(string yymm);
        void Serialize(string yymm);
        void CreateReportForm(string form);
        string GetCurrentTheme();
        string GetReportInfo();
        void SaveScan(string inUri,int num);
        void DeleteScan(int num);
        void ChangeStatus(ReportStatus status);
        void ChangeDataSource(DataSource datasource);
        void DeserializeReport(string yymm);
        void MapForAutoFill(AbstractReport report);
        void CallculateCells();
     
    }
}