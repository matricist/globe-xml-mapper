using System.Collections.Generic;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    public class Mapping_1_1_1_2 : MappingBase
    {
        public Mapping_1_1_1_2() : base("mapping_1.1~1.2.json") { }

        public override void Map(IXLWorksheet ws, Globe.GlobeOecd globe, List<string> errors, string fileName)
        {
            globe.GlobeBody.FilingInfo ??= new Globe.GlobeBodyTypeFilingInfo();
            var fi = globe.GlobeBody.FilingInfo;
            fi.FilingCe ??= new Globe.FilingInfoFilingCe();
            fi.AccountingInfo ??= new Globe.FilingInfoAccountingInfo();
            fi.Period ??= new Globe.FilingInfoPeriod();
            globe.GlobeBody.GeneralSection ??= new Globe.GlobeBodyTypeGeneralSection();
            globe.GlobeBody.GeneralSection.CorporateStructure ??= new Globe.CorporateStructureType();

            foreach (var (_, section) in Mapping.Sections)
            {
                foreach (var m in section.Mappings)
                {
                    ForEachValue(ws, m, fileName, errors, val =>
                    {
                        switch (m.Target)
                        {
                            case "FilingInfo.Period.Start":
                                if (TryParseDate(val, out var sd)) fi.Period.Start = sd;
                                else errors.Add($"[{fileName}] 셀 {m.Cell}: 날짜 변환 실패 '{val}'");
                                break;
                            case "FilingInfo.Period.End":
                                if (TryParseDate(val, out var ed)) fi.Period.End = ed;
                                else errors.Add($"[{fileName}] 셀 {m.Cell}: 날짜 변환 실패 '{val}'");
                                break;
                            case "FilingInfo.FilingCe.Name": fi.FilingCe.Name = val; break;
                            case "FilingInfo.FilingCe.KName": fi.FilingCe.KName = val; break;
                            case "FilingInfo.FilingCe.Tin.Value":
                                fi.FilingCe.Tin ??= new Globe.TinType();
                                fi.FilingCe.Tin.Value = val; break;
                            case "FilingInfo.FilingCe.ResCountryCode":
                                SetEnum<Globe.CountryCodeType>(val, v => fi.FilingCe.ResCountryCode = v, errors, fileName, m); break;
                            case "FilingInfo.FilingCe.Role":
                                SetEnum<Globe.FilingCeRoleEnumType>(val, v => fi.FilingCe.Role = v, errors, fileName, m); break;
                            case "FilingInfo.NameMne": fi.NameMne = val; break;
                            case "MessageSpec.MessageTypeIndic":
                                SetEnum<Globe.MessageTypeIndicEnumType>(val, v => globe.MessageSpec.MessageTypeIndic = v, errors, fileName, m); break;
                            case "FilingInfo.AccountingInfo.CfSofUpe":
                                SetEnum<Globe.FilingCeCofUpeEnumType>(val, v => fi.AccountingInfo.CfSofUpe = v, errors, fileName, m); break;
                            case "FilingInfo.AccountingInfo.Fas": fi.AccountingInfo.Fas = val; break;
                            case "FilingInfo.AccountingInfo.Currency":
                                SetEnum<Globe.CurrCodeType>(val, v => fi.AccountingInfo.Currency = v, errors, fileName, m); break;
                            case "GeneralSection.RecJurCode":
                                SetEnum<Globe.CountryCodeType>(val, v => globe.GlobeBody.GeneralSection.RecJurCode.Add(v), errors, fileName, m); break;
                            default:
                                errors.Add($"[{fileName}] 셀 {m.Cell}: 알 수 없는 매핑 대상 '{m.Target}'"); break;
                        }
                    });
                }
            }
        }
    }
}
