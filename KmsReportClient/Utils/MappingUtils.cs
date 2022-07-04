using System.Collections.Generic;
using System.Linq;
using KmsReportClient.Report;

namespace KmsReportClient.Utils
{
    static class MappingUtils
    {
        public static void MapFromExternal(AbstractReport dst, External.AbstractReport src)
        {
            if (dst is Report294 && src is External.Report294)
            {
                MapExternalToReport294((Report294)dst, (External.Report294)src);
            }
            else if (dst is Report262 && src is External.Report262)
            {
                MapExternalToReport262((Report262)dst, (External.Report262)src);
            }
            else if (dst is ReportIizl && src is External.ReportIizl)
            {
                MapExternalToReportIizl((ReportIizl)dst, (External.ReportIizl)src);
            }
            else if (dst is ReportPg && src is External.ReportPg)
            {
                MapExternalToReportPg((ReportPg)dst, (External.ReportPg)src);
            }
        }

        private static void MapAbstractToExternal(AbstractReport dst, External.AbstractReport src)
        {
            dst.Created = src.Created;
            dst.RefuseDate = src.DateEditCo;
            dst.Version = src.Version;
            dst.DateIsDone = src.DateIsDone;
            dst.DateToCo = src.DateToCo;
            dst.IdEmployeeUpd = src.IdEmployeeUpd;
            dst.IsDone = src.IsDone;
            dst.Submit = src.Submit;
            dst.Refuse = src.Refuse;
            dst.Scan = src.Scan;
            dst.Updated = src.Updated;
            dst.RefuseUser = src.UserEditCo;
            dst.UserSubmit = src.UserSubmit;
            dst.UserToCo = src.UserToCo;
            dst.IdEmployee = src.IdEmployee;
        }

        public static External.AbstractReport MapToExternalAbstractReport(AbstractReport src) =>
            new External.AbstractReport
            {
                Id = src.Id,
                Name = src.Name,
                OldTheme = src.OldTheme,
                SerializeName = src.SerializeName,
                SmallName = src.SmallName,
                Yymm = src.Yymm,
                Created = src.Created,
                DateEditCo = src.RefuseDate,
                DateIsDone = src.DateIsDone,
                DateToCo = src.DateToCo,
                IdEmployeeUpd = src.IdEmployeeUpd,
                IsDone = src.IsDone,
                Scan = src.Scan,
                Updated = src.Updated,
                UserEditCo = src.RefuseUser,
                UserSubmit = src.UserSubmit,
                UserToCo = src.UserToCo,
                Version = src.Version
            };

        private static void MapExternalToReportPg(ReportPg dst, External.ReportPg src)
        {
            MapAbstractToExternal(dst, src);

            foreach (var theme in src.ReportDataList)
            {
                var currentReportData = dst.ReportDataList.SingleOrDefault(x => x.Theme == theme.Theme);
                if (currentReportData != null)
                {
                    currentReportData.Data = new List<ReportPgDataDto>();

                    foreach (var data in theme.Data)
                    {
                        var dataDto = new ReportPgDataDto
                        {
                            Code = data.Code,
                            CountSmo = data.CountSmo,
                            CountSmoAnother = data.CountSmoAnother,
                            CountInsured = data.CountInsured,
                            CountInsuredRepresentative = data.CountInsuredRepresentative,
                            CountTfoms = data.CountTfoms,
                            CountProsecutor = data.CountProsecutor,
                            CountOutOfSmo = data.CountOutOfSmo,
                            CountAmbulatory = data.CountAmbulatory,
                            CountDs = data.CountDs,
                            CountDsVmp = data.CountDsVmp,
                            CountStac = data.CountStac,
                            CountStacVmp = data.CountStacVmp,
                            CountOutOfSmoAnother = data.CountOutOfSmoAnother,
                            CountAmbulatoryAnother = data.CountAmbulatoryAnother,
                            CountDsAnother = data.CountDsAnother,
                            CountDsVmpAnother = data.CountDsVmpAnother,
                            CountStacAnother = data.CountStacAnother,
                            CountStacVmpAnother = data.CountStacVmpAnother,
                        };

                        currentReportData.Data.Add(dataDto);
                    }

                }
            }
        }

        public static External.ReportPg MapReportPgToExternal(ReportPg src)
        {
            var dst = new External.ReportPg
            {
                Id = src.Id,
                Name = src.Name,
                OldTheme = src.OldTheme,
                SerializeName = src.SerializeName,
                SmallName = src.SmallName,
                Yymm = src.Yymm,
                Created = src.Created,
                DateEditCo = src.RefuseDate,
                DateIsDone = src.DateIsDone,
                DateToCo = src.DateToCo,
                IdEmployeeUpd = src.IdEmployeeUpd,
                IsDone = src.IsDone,
                Scan = src.Scan,
                Updated = src.Updated,
                UserEditCo = src.RefuseUser,
                UserSubmit = src.UserSubmit,
                UserToCo = src.UserToCo,
                Version = src.Version
            };

            var dataList = new List<External.ReportPgDto>();
            foreach (var srcTheme in src.ReportDataList)
            {
                var dstDtos = new List<External.ReportPgDataDto>();
                foreach (var srcDto in srcTheme.Data)
                {
                    var dstDto = new External.ReportPgDataDto
                    {
                        Code = srcDto.Code,
                        CountSmo = srcDto.CountSmo,
                        CountSmoAnother = srcDto.CountSmoAnother,
                        CountInsured = srcDto.CountInsured,
                        CountInsuredRepresentative = srcDto.CountInsuredRepresentative,
                        CountTfoms = srcDto.CountTfoms,
                        CountProsecutor = srcDto.CountProsecutor,
                        CountOutOfSmo = srcDto.CountOutOfSmo,
                        CountAmbulatory = srcDto.CountAmbulatory,
                        CountDs = srcDto.CountDs,
                        CountDsVmp = srcDto.CountDsVmp,
                        CountStac = srcDto.CountStac,
                        CountStacVmp = srcDto.CountStacVmp,
                        CountOutOfSmoAnother = srcDto.CountOutOfSmoAnother,
                        CountAmbulatoryAnother = srcDto.CountAmbulatoryAnother,
                        CountDsAnother = srcDto.CountDsAnother,
                        CountDsVmpAnother = srcDto.CountDsVmpAnother,
                        CountStacAnother = srcDto.CountStacAnother,
                        CountStacVmpAnother = srcDto.CountStacVmpAnother
                    };
                    dstDtos.Add(dstDto);
                }

                var dstData = new External.ReportPgDto
                {
                    Theme = srcTheme.Theme,
                    Data = dstDtos.ToArray(),
                };
                dataList.Add(dstData);
            }
            dst.ReportDataList = dataList.ToArray();

            return dst;
        }

        private static void MapExternalToReportIizl(ReportIizl dst, External.ReportIizl src)
        {
            MapAbstractToExternal(dst, src);

            foreach (var theme in src.ReportDataList)
            {
                var currentReportData = dst.ReportDataList.SingleOrDefault(x => x.Theme == theme.Theme);
                if (currentReportData != null)
                {
                    currentReportData.TotalPersFirst = theme.TotalPersFirst;
                    currentReportData.TotalPersRepeat = theme.TotalPersRepeat;
                    currentReportData.Data = new List<ReportIizlDataDto>();

                    foreach (var data in theme.Data)
                    {
                        var dataDto = new ReportIizlDataDto
                        {
                            TotalCost = data.TotalCost,
                            CountPersRepeat = data.CountPersRepeat,
                            CountMessages = data.CountMessages,
                            AccountingDocument = data.AccountingDocument,
                            CountPersFirst = data.CountPersFirst,
                            Code = data.Code
                        };

                        currentReportData.Data.Add(dataDto);
                    }

                }
            }
        }

        public static External.ReportIizl MapReportIizlToExternal(ReportIizl src)
        {
            var dst = new External.ReportIizl
            {
                Id = src.Id,
                Name = src.Name,
                OldTheme = src.OldTheme,
                SerializeName = src.SerializeName,
                SmallName = src.SmallName,
                Yymm = src.Yymm,
                Created = src.Created,
                DateEditCo = src.RefuseDate,
                DateIsDone = src.DateIsDone,
                DateToCo = src.DateToCo,
                IdEmployeeUpd = src.IdEmployeeUpd,
                IsDone = src.IsDone,
                Scan = src.Scan,
                Updated = src.Updated,
                UserEditCo = src.RefuseUser,
                UserSubmit = src.UserSubmit,
                UserToCo = src.UserToCo,
                Version = src.Version
            };

            var dataList = new List<External.ReportIizlDto>();
            foreach (var srcTheme in src.ReportDataList)
            {
                var dstDtos = new List<External.ReportIizlDataDto>();
                foreach (var srcDto in srcTheme.Data)
                {
                    var dstDto = new External.ReportIizlDataDto
                    {
                        AccountingDocument = srcDto.AccountingDocument,
                        Code = srcDto.Code,
                        CountMessages = srcDto.CountMessages,
                        CountPersFirst = srcDto.CountPersFirst,
                        CountPersRepeat = srcDto.CountPersRepeat,
                        TotalCost = srcDto.TotalCost
                    };
                    dstDtos.Add(dstDto);
                }

                var dstData = new External.ReportIizlDto
                {
                    Theme = srcTheme.Theme,
                    Data = dstDtos.ToArray(),
                    TotalPersFirst = srcTheme.TotalPersFirst,
                    TotalPersRepeat = srcTheme.TotalPersRepeat
                };
                dataList.Add(dstData);
            }
            dst.ReportDataList = dataList.ToArray();

            return dst;
        }

        private static void MapExternalToReport262(Report262 dst, External.Report262 src)
        {
            MapAbstractToExternal(dst, src);

            var t3 = dst.ReportDataList.SingleOrDefault(x => x.Theme == "Таблица 3");
            t3?.Table3.Clear();
            var d1 = dst.ReportDataList.SingleOrDefault(x => x.Theme == "Таблица 1");
            d1?.Data.Clear();
            var d2 = dst.ReportDataList.SingleOrDefault(x => x.Theme == "Таблица 2");
            d2?.Data.Clear();

            foreach (var srcData in src.ReportDataList)
            {
                var dstData = dst.ReportDataList.SingleOrDefault(x => x.Theme == srcData.Theme);
                if (dstData == null)
                {
                    continue;
                }

                string theme = srcData.Theme;
                if (theme == "Таблица 3")
                {
                    foreach (var table3 in srcData.Table3)
                    {
                        var table3Data = new Report262Table3Data
                        {
                            CountChannelAnother = table3.CountChannelAnother,
                            CountChannelAnotherChild = table3.CountChannelAnotherChild,
                            CountChannelPhone = table3.CountChannelPhone,
                            CountChannelPhoneChild = table3.CountChannelPhoneChild,
                            CountChannelSp = table3.CountChannelSp,
                            CountChannelSpChild = table3.CountChannelSpChild,
                            CountChannelTerminal = table3.CountChannelTerminal,
                            CountChannelTerminalChild = table3.CountChannelTerminalChild,
                            CountUnit = table3.CountUnit,
                            CountUnitChild = table3.CountUnitChild,
                            CountUnitWithSp = table3.CountUnitWithSp,
                            CountUnitWithSpChild = table3.CountUnitWithSpChild,
                            Mo = table3.Mo
                        };
                        dstData.Table3.Add(table3Data);
                    }
                }
                else
                {
                    Report262DataDto dataDto;
                    if (srcData.Data != null && srcData.Data.Count() == 1)
                    {
                        var data = srcData.Data[0];
                        dataDto = new Report262DataDto
                        {
                            CountAddress = data.CountAddress,
                            CountAnother = data.CountAnother,
                            CountEmail = data.CountEmail,
                            CountMessengers = data.CountMessengers,
                            CountPhone = data.CountPhone,
                            CountPost = data.CountPost,
                            CountPpl = data.CountPpl,
                            CountPplFull = data.CountPplFull,
                            CountSms = data.CountSms,
                            RowNum = data.RowNum
                        };
                    }
                    else
                    {
                        dataDto = new Report262DataDto
                        {
                            CountAddress = 0,
                            CountAnother = 0,
                            CountEmail = 0,
                            CountMessengers = 0,
                            CountPhone = 0,
                            CountPost = 0,
                            CountPpl = 0,
                            CountPplFull = 0,
                            CountSms = 0,
                            RowNum = 1
                        };
                    }
                    dstData.Data.Add(dataDto);
                }
            }
        }

        public static External.Report262 MapReport262ToExternal(Report262 src)
        {
            var dst = new External.Report262
            {
                Id = src.Id,
                Name = src.Name,
                OldTheme = src.OldTheme,
                SerializeName = src.SerializeName,
                SmallName = src.SmallName,
                Yymm = src.Yymm,
                Created = src.Created,
                DateEditCo = src.RefuseDate,
                DateIsDone = src.DateIsDone,
                DateToCo = src.DateToCo,
                IdEmployeeUpd = src.IdEmployeeUpd,
                IsDone = src.IsDone,
                Scan = src.Scan,
                Updated = src.Updated,
                UserEditCo = src.RefuseUser,
                UserSubmit = src.UserSubmit,
                UserToCo = src.UserToCo,
                Version = src.Version
            };

            var dataList = new List<External.Report262Dto>();
            foreach (var srcTheme in src.ReportDataList)
            {
                var dstDtos = new List<External.Report262DataDto>();
                foreach (var srcDto in srcTheme.Data)
                {
                    var dstDto = new External.Report262DataDto
                    {
                        CountAddress = srcDto.CountAddress,
                        CountAnother = srcDto.CountAnother,
                        CountEmail = srcDto.CountEmail,
                        CountMessengers = srcDto.CountMessengers,
                        CountPhone = srcDto.CountPhone,
                        CountPost = srcDto.CountPost,
                        CountPpl = srcDto.CountPpl,
                        CountPplFull = srcDto.CountPplFull,
                        CountSms = srcDto.CountSms,
                        RowNum = srcDto.RowNum
                    };
                    dstDtos.Add(dstDto);
                }

                var dstTable3 = new List<External.Report262Table3Data>();
                foreach (var srcT in srcTheme.Table3)
                {
                    var dstT = new External.Report262Table3Data
                    {
                        CountChannelAnother = srcT.CountChannelAnother,
                        CountChannelAnotherChild = srcT.CountChannelAnotherChild,
                        CountChannelPhone = srcT.CountChannelPhone,
                        CountChannelPhoneChild = srcT.CountChannelPhoneChild,
                        CountChannelSp = srcT.CountChannelSp,
                        CountChannelSpChild = srcT.CountChannelSpChild,
                        CountChannelTerminal = srcT.CountChannelTerminal,
                        CountChannelTerminalChild = srcT.CountChannelTerminalChild,
                        CountUnit = srcT.CountUnit,
                        CountUnitChild = srcT.CountUnitChild,
                        CountUnitWithSp = srcT.CountUnitWithSp,
                        CountUnitWithSpChild = srcT.CountUnitWithSpChild,
                        Mo = srcT.Mo
                    };
                    dstTable3.Add(dstT);
                }

                var dstData = new External.Report262Dto
                {
                    Theme = srcTheme.Theme,
                    Data = dstDtos.ToArray(),
                    Table3 = dstTable3.ToArray()
                };
                dataList.Add(dstData);
            }
            dst.ReportDataList = dataList.ToArray();
            return dst;
        }

        private static void MapExternalToReport294(Report294 dst, External.Report294 src)
        {
            MapAbstractToExternal(dst, src);

            foreach (var theme in src.ReportDataList)
            {
                var currentReportData = dst.ReportDataList.SingleOrDefault(x => x.Theme == theme.Theme);
                if (currentReportData != null)
                {
                    currentReportData.Data = new List<Report294DataDto>();

                    foreach (var data in theme.Data)
                    {
                        var dataDto = new Report294DataDto
                        {
                            CountAddress = data.CountAddress,
                            CountAnother = data.CountAnother,
                            CountAnotherDisease = data.CountAnotherDisease,
                            CountBloodDisease = data.CountBloodDisease,
                            CountBronchoDisease = data.CountBronchoDisease,
                            CountEmail = data.CountEmail,
                            CountEndocrineDisease = data.CountEndocrineDisease,
                            CountMessangers = data.CountMessengers,
                            CountOncologicalDisease = data.CountOncologicalDisease,
                            CountPhone = data.CountPhone,
                            CountPost = data.CountPost,
                            CountPpl = data.CountPpl,
                            CountSms = data.CountSms,
                            RowNum = data.RowNum
                        };
                        currentReportData.Data.Add(dataDto);
                    }
                }
            }
        }

        public static External.Report294 MapReport294ToExternal(Report294 src)
        {
            var dst = new External.Report294
            {
                Id = src.Id,
                Name = src.Name,
                OldTheme = src.OldTheme,
                SerializeName = src.SerializeName,
                SmallName = src.SmallName,
                Yymm = src.Yymm,
                Created = src.Created,
                DateEditCo = src.RefuseDate,
                DateIsDone = src.DateIsDone,
                DateToCo = src.DateToCo,
                IdEmployeeUpd = src.IdEmployeeUpd,
                IsDone = src.IsDone,
                Scan = src.Scan,
                Updated = src.Updated,
                UserEditCo = src.RefuseUser,
                UserSubmit = src.UserSubmit,
                UserToCo = src.UserToCo,
                Version = src.Version
            };

            var dataList = new List<External.Report294Dto>();
            foreach (var srcTheme in src.ReportDataList)
            {
                var dstDtos = new List<External.Report294DataDto>();
                foreach (var srcDto in srcTheme.Data)
                {
                    var dstDto = new External.Report294DataDto
                    {
                        CountAddress = srcDto.CountAddress,
                        CountAnother = srcDto.CountAnother,
                        RowNum = srcDto.RowNum,
                        CountSms = srcDto.CountSms,
                        CountPpl = srcDto.CountPpl,
                        CountPost = srcDto.CountPost,
                        CountAnotherDisease = srcDto.CountAnotherDisease,
                        CountBloodDisease = srcDto.CountBloodDisease,
                        CountBronchoDisease = srcDto.CountBronchoDisease,
                        CountEmail = srcDto.CountEmail,
                        CountEndocrineDisease = srcDto.CountEndocrineDisease,
                        CountMessengers = srcDto.CountMessangers,
                        CountOncologicalDisease = srcDto.CountOncologicalDisease,
                        CountPhone = srcDto.CountPhone
                    };
                    dstDtos.Add(dstDto);
                }

                var dstData = new External.Report294Dto
                {
                    Theme = srcTheme.Theme,
                    Data = dstDtos.ToArray()
                };
                dataList.Add(dstData);
            }
            dst.ReportDataList = dataList.ToArray();

            return dst;
        }

        public static List<KmsReportDictionary> MapKmsReportDictionary(External.KmsReportDictionary[] srcList)
        {
            var dstlist = new List<KmsReportDictionary>();
            foreach (var src in srcList)
            {
                var dst = new KmsReportDictionary
                {
                    ForeignKey = src.ForeignKey,
                    Key = src.Key,
                    Value = src.Value
                };
                dstlist.Add(dst);
            }
            return dstlist;
        }
    }
}