using System;
using System.Collections.Generic;
using KmsReportClient.External;

namespace KmsReportClient.Global
{
    [Serializable]
    static class CurrentUser
    {
        public static string Filial;
        public static string Region;
        public static string FilialCode;

        public static int IdUser;
        public static string UserName;
        public static string Phone;
        public static string Email;

        public static string Director;
        public static string DirectorPosition;
        public static string DirectorPhone;

        public static bool IsMain;

        public static List<KmsReportDictionary> Regions;
        public static List<KmsReportDictionary> Users;
        public static List<KmsReportDictionary> ReportTypes;
        
       
    }
}