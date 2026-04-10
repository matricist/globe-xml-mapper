using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace GlobeMapper.Services
{
    /// <summary>
    /// GIR XML 에러코드 매뉴얼 기반 유효성 검증
    /// </summary>
    public static class ValidationUtil
    {
        public static List<string> Validate(Globe.GlobeOecd globe)
        {
            var errors = new List<string>();

            ValidateMessageSpec(globe, errors);
            ValidateDocSpec(globe, errors);
            ValidateFilingInfo(globe, errors);
            ValidateRecJurCode(globe, errors);
            ValidateUpe(globe, errors);
            ValidateCe(globe, errors);
            ValidateRulesConsistency(globe, errors);
            ValidateTin(globe, errors);
            ValidateSummary(globe, errors);

            return errors;
        }

        #region MessageSpec (60001, 60003)

        private static void ValidateMessageSpec(Globe.GlobeOecd globe, List<string> errors)
        {
            var spec = globe.MessageSpec;
            if (spec == null) return;

            // 60001: MessageRefId 형식 [발신국가코드][보고기간][수신국가코드][고유식별번호]
            if (string.IsNullOrEmpty(spec.MessageRefId))
            {
                errors.Add("[60001] [MessageSpec] MessageRefId가 비어 있습니다.");
            }

            // 60003: ReportingPeriod YYYY ≤ 현재연도
            if (spec.ReportingPeriod != default && spec.ReportingPeriod.Year > DateTime.Now.Year)
            {
                errors.Add($"[60003] [MessageSpec] ReportingPeriod 연도({spec.ReportingPeriod.Year})가 현재 연도({DateTime.Now.Year})보다 큽니다.");
            }
        }

        #endregion

        #region DocSpec (60004, 60011, 60012, 60017)

        private static void ValidateDocSpec(Globe.GlobeOecd globe, List<string> errors)
        {
            var body = globe.GlobeBody;
            if (body == null) return;

            // 60016: OECD0(재제출) FilingInfo이면 GeneralSection은 OECD1 불가
            if (body.FilingInfo?.DocSpec?.DocTypeIndic == Globe.OecdDocTypeIndicEnumType.Oecd0
                && body.GeneralSection?.DocSpec?.DocTypeIndic == Globe.OecdDocTypeIndicEnumType.Oecd1)
            {
                errors.Add("[60016] [DocSpec] FilingInfo DocTypeIndic이 OECD0(재제출)인 경우 GeneralSection은 OECD1(신규)이 될 수 없습니다.");
            }

            // 60013: OECD0(재제출)은 FilingInfo에만 사용 가능
            if (body.GeneralSection?.DocSpec?.DocTypeIndic == Globe.OecdDocTypeIndicEnumType.Oecd0)
                errors.Add("[60013] [GeneralSection/DocSpec] OECD0(재제출)은 FilingInfo에만 사용 가능합니다. GeneralSection에 OECD0을 사용할 수 없습니다.");

            // 60017: FilingInfo DocTypeIndic=OECD1이면 GeneralSection 필수
            if (body.FilingInfo?.DocSpec?.DocTypeIndic == Globe.OecdDocTypeIndicEnumType.Oecd1
                && body.GeneralSection == null)
            {
                errors.Add("[60017] [FilingInfo/DocSpec] DocTypeIndic이 OECD1이면 GeneralSection이 필수입니다.");
            }

            // FilingInfo DocSpec 검증
            ValidateSingleDocSpec(body.FilingInfo?.DocSpec, "FilingInfo", errors);

            // GeneralSection DocSpec 검증
            ValidateSingleDocSpec(body.GeneralSection?.DocSpec, "GeneralSection", errors);

            // 60004: 신규(OECD1)와 정정(OECD2/OECD3) 혼합 불가
            var docTypes = new List<Globe.OecdDocTypeIndicEnumType?>();
            docTypes.Add(body.FilingInfo?.DocSpec?.DocTypeIndic);
            docTypes.Add(body.GeneralSection?.DocSpec?.DocTypeIndic);
            var distinct = docTypes.Where(d => d.HasValue).Select(d => d.Value).Distinct().ToList();
            if (distinct.Contains(Globe.OecdDocTypeIndicEnumType.Oecd1)
                && (distinct.Contains(Globe.OecdDocTypeIndicEnumType.Oecd2) || distinct.Contains(Globe.OecdDocTypeIndicEnumType.Oecd3)))
            {
                errors.Add("[60004] [DocSpec] 하나의 메시지에 신규(OECD1)와 정정(OECD2/OECD3)을 혼합할 수 없습니다.");
            }
        }

        private static void ValidateSingleDocSpec(Globe.DocSpecType docSpec, string section, List<string> errors)
        {
            if (docSpec == null)
            {
                errors.Add($"[60011] [{section}/DocSpec] DocSpec이 없습니다.");
                return;
            }

            // 60011: DocRefId 형식 [발신관할권국가코드][보고연도][고유식별번호]
            if (string.IsNullOrEmpty(docSpec.DocRefId))
            {
                errors.Add($"[60011] [{section}/DocSpec] DocRefId가 비어 있습니다.");
            }

            // 60012: OECD1/OECD0이면 CorrDocRefId 생략
            if ((docSpec.DocTypeIndic == Globe.OecdDocTypeIndicEnumType.Oecd1
                 || docSpec.DocTypeIndic == Globe.OecdDocTypeIndicEnumType.Oecd0)
                && !string.IsNullOrEmpty(docSpec.CorrDocRefId))
            {
                errors.Add($"[60012] [{section}/DocSpec] DocTypeIndic이 OECD1/OECD0인 경우 CorrDocRefId는 생략되어야 합니다.");
            }

            // 60015: OECD2/OECD3이면 CorrDocRefId 필수
            if ((docSpec.DocTypeIndic == Globe.OecdDocTypeIndicEnumType.Oecd2
                 || docSpec.DocTypeIndic == Globe.OecdDocTypeIndicEnumType.Oecd3)
                && string.IsNullOrEmpty(docSpec.CorrDocRefId))
            {
                errors.Add($"[60015] [{section}/DocSpec] DocTypeIndic이 OECD2/OECD3인 경우 CorrDocRefId가 필수입니다.");
            }
        }

        #endregion

        #region FilingInfo (60020, 60021, 60022, 60023)

        private static void ValidateFilingInfo(Globe.GlobeOecd globe, List<string> errors)
        {
            var filingInfo = globe.GlobeBody?.FilingInfo;
            if (filingInfo == null) return;

            var period = filingInfo.Period;
            if (period != null)
            {
                // 60020: Start ≤ End
                if (period.Start != default && period.End != default && period.Start > period.End)
                {
                    errors.Add($"[60020] [1.1~1.2 신고기본정보/Period] 기간 시작일({period.Start:yyyy-MM-dd})이 종료일({period.End:yyyy-MM-dd})보다 늦습니다.");
                }

                // 60021: Period End ≤ ReportingPeriod
                if (period.End != default && globe.MessageSpec?.ReportingPeriod != default
                    && period.End > globe.MessageSpec.ReportingPeriod)
                {
                    errors.Add($"[60021] [1.1~1.2 신고기본정보/Period] 기간 종료일({period.End:yyyy-MM-dd})이 ReportingPeriod({globe.MessageSpec.ReportingPeriod:yyyy-MM-dd})보다 늦습니다.");
                }
            }

            // 60023: FilingCE ResCountryCode = TransmittingCountry
            if (filingInfo.FilingCe != null && globe.MessageSpec != null
                && filingInfo.FilingCe.ResCountryCode != globe.MessageSpec.TransmittingCountry)
            {
                errors.Add($"[60023] [1.1~1.2 신고기본정보/FilingCE] 소재지국({filingInfo.FilingCe.ResCountryCode})이 TransmittingCountry({globe.MessageSpec.TransmittingCountry})와 불일치합니다.");
            }

            // 60022: Role=GIR401이면 FilingCE TIN이 UPE TIN 중 하나와 일치해야 함
            if (filingInfo.FilingCe?.Role == Globe.FilingCeRoleEnumType.Gir401
                && filingInfo.FilingCe?.Tin != null)
            {
                var filingTin = filingInfo.FilingCe.Tin.Value;
                var upeTins = GetUpeTins(globe);

                if (!string.IsNullOrEmpty(filingTin) && upeTins.Count > 0
                    && !upeTins.Contains(filingTin))
                {
                    errors.Add($"[60022] [1.1~1.2 신고기본정보/FilingCE] Role=GIR401(UPE)인데 FilingCE TIN({filingTin})이 UPE TIN({string.Join(", ", upeTins)}) 중 어느 것과도 일치하지 않습니다.");
                }
            }
        }

        private static List<string> GetUpeTins(Globe.GlobeOecd globe)
        {
            var tins = new List<string>();
            var corpStructure = globe.GlobeBody?.GeneralSection?.CorporateStructure;
            if (corpStructure == null) return tins;

            foreach (var upe in corpStructure.Upe)
            {
                if (upe.OtherUpe?.Id?.Tin != null)
                    tins.AddRange(upe.OtherUpe.Id.Tin.Select(t => t.Value));
                if (upe.ExcludedUpe?.Id?.Tin != null)
                    tins.AddRange(upe.ExcludedUpe.Id.Tin.Select(t => t.Value));
            }

            return tins.Where(t => !string.IsNullOrEmpty(t)).ToList();
        }

        #endregion

        #region RecJurCode (60018)

        private static void ValidateRecJurCode(Globe.GlobeOecd globe, List<string> errors)
        {
            var recJurCodes = globe.GlobeBody?.GeneralSection?.RecJurCode;
            if (recJurCodes == null || recJurCodes.Count == 0) return;

            // 60019: Role GIR403/404/405이면 로컬 제출 (교환 불가)
            var role = globe.GlobeBody?.FilingInfo?.FilingCe?.Role;
            if (role == Globe.FilingCeRoleEnumType.Gir403
                || role == Globe.FilingCeRoleEnumType.Gir404
                || role == Globe.FilingCeRoleEnumType.Gir405)
            {
                errors.Add($"[60019] [1.1~1.2 신고기본정보/RecJurCode] FilingCE Role({role})은 로컬 제출(Local Lodgement)로, 자동교환 대상이 아닙니다.");
            }

            // 60018: RecJurCode 중 하나가 ReceivingCountry와 일치해야 함
            var recvCountry = globe.MessageSpec?.ReceivingCountry;
            if (recvCountry != null && !recJurCodes.Contains(recvCountry.Value))
            {
                errors.Add($"[60018] [1.1~1.2 신고기본정보/RecJurCode] ReceivingCountry({recvCountry})가 RecJurCode에 포함되어 있지 않습니다.");
            }
        }

        #endregion

        #region UPE (70009, 70010)

        private static void ValidateUpe(Globe.GlobeOecd globe, List<string> errors)
        {
            var corpStructure = globe.GlobeBody?.GeneralSection?.CorporateStructure;
            if (corpStructure == null) return;

            // 70009: UPE GlobeStatus에 허용되지 않는 값
            var disallowed = new HashSet<Globe.IdTypeGloBeStatusEnumType>
            {
                Globe.IdTypeGloBeStatusEnumType.Gir305, // PE
                Globe.IdTypeGloBeStatusEnumType.Gir307, // Minority-Owned Parent
                Globe.IdTypeGloBeStatusEnumType.Gir308, // Minority-Owned Subsidiary
                Globe.IdTypeGloBeStatusEnumType.Gir309,
                Globe.IdTypeGloBeStatusEnumType.Gir312,
                Globe.IdTypeGloBeStatusEnumType.Gir313,
                Globe.IdTypeGloBeStatusEnumType.Gir314,
                Globe.IdTypeGloBeStatusEnumType.Gir315,
                Globe.IdTypeGloBeStatusEnumType.Gir317,
                Globe.IdTypeGloBeStatusEnumType.Gir318,
            };

            foreach (var upe in corpStructure.Upe)
            {
                var statuses = new List<Globe.IdTypeGloBeStatusEnumType>();
                if (upe.OtherUpe?.Id?.GlobeStatus != null)
                    statuses.AddRange(upe.OtherUpe.Id.GlobeStatus);
                if (upe.ExcludedUpe?.Id?.GlobeStatus != null)
                    statuses.AddRange(upe.ExcludedUpe.Id.GlobeStatus);

                var upeLabel = upe.OtherUpe?.Id?.Name ?? upe.ExcludedUpe?.Id?.Name ?? "(이름없음)";
                foreach (var status in statuses)
                {
                    if (disallowed.Contains(status))
                    {
                        errors.Add($"[70009] [1.3.1 최종모기업/'{upeLabel}'] GlobeStatus에 허용되지 않는 값({status})이 포함되어 있습니다.");
                    }
                }

                // 70010: OtherUPE ResCountryCode는 하나만 허용
                if (upe.OtherUpe?.Id?.ResCountryCode?.Count > 1)
                {
                    errors.Add($"[70010] [1.3.1 최종모기업/'{upeLabel}'] ResCountryCode에 {upe.OtherUpe.Id.ResCountryCode.Count}개 값이 있습니다. 하나만 허용됩니다.");
                }
            }
        }

        #endregion

        #region CE (70011~70032)

        private static void ValidateCe(Globe.GlobeOecd globe, List<string> errors)
        {
            var corpStructure = globe.GlobeBody?.GeneralSection?.CorporateStructure;
            if (corpStructure == null) return;

            foreach (var ce in corpStructure.Ce)
            {
                var ceName = ce.Id?.Name ?? "(이름없음)";
                var loc = $"1.3.2.1 구성기업/'{ceName}'";

                // 70011: CE ResCountryCode 하나만 허용
                if (ce.Id?.ResCountryCode?.Count > 1)
                {
                    errors.Add($"[70011] [{loc}] ResCountryCode에 {ce.Id.ResCountryCode.Count}개 값이 있습니다. 하나만 허용됩니다.");
                }

                // 70013: GIR313(JV)과 GIR314(JV Subsidiary) 동시 불가
                var statuses = ce.Id?.GlobeStatus;
                if (statuses != null)
                {
                    if (statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir313)
                        && statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir314))
                    {
                        errors.Add($"[70013] [{loc}] GIR313(JV)과 GIR314(JV Subsidiary)가 동시에 보고되었습니다.");
                    }

                    // 70014: GIR307과 GIR308 동시 불가
                    if (statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir307)
                        && statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir308))
                    {
                        errors.Add($"[70014] [{loc}] GIR307(소수지분모기업)과 GIR308(소수지분 자회사)가 동시에 보고되었습니다.");
                    }

                    // 70018: GIR305(PE)와 GIR306(Main Entity) 동시 불가
                    if (statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir305)
                        && statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir306))
                    {
                        errors.Add($"[70018] [{loc}] GIR305(고정사업장)과 GIR306(본점)이 동시에 보고되었습니다.");
                    }

                    // 70020: GIR316/GIR318은 유일한 값이어야 함
                    if ((statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir316)
                         || statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir318))
                        && statuses.Count > 1)
                    {
                        errors.Add($"[70020] [{loc}] GIR316/GIR318은 GlobeStatus에서 유일한 값이어야 합니다.");
                    }

                    // 70016: GIR307 있으면 GIR309도 필요
                    if (statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir307)
                        && !statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir309))
                        errors.Add($"[70016] [{loc}] GIR307(소수지분모기업)이 있으면 GIR309도 함께 보고해야 합니다.");

                    // 70017: GIR308 있으면 GIR309도 필요
                    if (statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir308)
                        && !statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir309))
                        errors.Add($"[70017] [{loc}] GIR308(소수지분 자회사)이 있으면 GIR309도 함께 보고해야 합니다.");

                    // 70021: GIR316/318이면 OwnershipChange 필수
                    if ((statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir316)
                         || statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir318))
                        && ce.OwnershipChange.Count == 0)
                        errors.Add($"[70021] [{loc}] GIR316 또는 GIR318이면 OwnershipChange(1.3.3 시트)가 있어야 합니다.");
                }

                // 70026: GIR305(PE)이면 OwnershipPercentage = 100%
                if (statuses != null && statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir305))
                {
                    foreach (var own in ce.Ownership)
                    {
                        if (own.OwnershipPercentage != 1.0m)
                        {
                            errors.Add($"[70026] [{loc}/Ownership] GIR305(고정사업장)이므로 OwnershipPercentage가 100%여야 합니다. (현재: {own.OwnershipPercentage:P0})");
                        }
                    }
                }

                // 70027: GIR318이면 OwnershipPercentage=0%, TIN=NOTIN, OwnershipType=GIR806
                if (statuses != null && statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir318))
                {
                    foreach (var own in ce.Ownership)
                    {
                        if (own.OwnershipPercentage != 0m)
                            errors.Add($"[70027] [{loc}/Ownership] GIR318이면 OwnershipPercentage=0%이어야 합니다. (현재: {own.OwnershipPercentage:P0})");
                        if (!string.Equals(own.Tin?.Value, "NOTIN", StringComparison.OrdinalIgnoreCase))
                            errors.Add($"[70027] [{loc}/Ownership] GIR318이면 Ownership/TIN='NOTIN'이어야 합니다.");
                        if (own.OwnershipType != Globe.OwnershipTypeEnumType.Gir806)
                            errors.Add($"[70027] [{loc}/Ownership] GIR318이면 OwnershipType=GIR806이어야 합니다.");
                    }
                }

                // 70028: GIR318이 아니면 OwnershipPercentage는 0% 불가
                if (statuses == null || !statuses.Contains(Globe.IdTypeGloBeStatusEnumType.Gir318))
                {
                    foreach (var own in ce.Ownership)
                    {
                        if (own.OwnershipPercentage == 0m)
                            errors.Add($"[70028] [{loc}/Ownership] GIR318이 아닌 경우 OwnershipPercentage는 0%가 될 수 없습니다.");
                    }
                }

                // 70032: QIIR 제공 시 Rules에 GIR201/GIR202 필수
                if (ce.Qiir != null && ce.Id?.Rules != null)
                {
                    if (!ce.Id.Rules.Contains(Globe.IdTypeRulesEnumType.Gir201)
                        && !ce.Id.Rules.Contains(Globe.IdTypeRulesEnumType.Gir202))
                    {
                        errors.Add($"[70032] [{loc}/QIIR] QIIR가 제공되었으나 Rules에 GIR201/GIR202가 포함되지 않았습니다.");
                    }
                }

                // 70033: QIIR/Exception/TIN은 다른 CE TIN과 일치해야
                if (ce.Qiir?.Exception?.Tin != null)
                {
                    var exTinVal = ce.Qiir.Exception.Tin.Value;
                    if (!string.IsNullOrEmpty(exTinVal)
                        && !string.Equals(exTinVal, "NOTIN", StringComparison.OrdinalIgnoreCase))
                    {
                        var allCeTins = corpStructure.Ce
                            .Where(c => c != ce)
                            .SelectMany(c => c.Id?.Tin ?? Enumerable.Empty<Globe.TinType>())
                            .Select(t => t.Value)
                            .ToHashSet();
                        if (!allCeTins.Contains(exTinVal))
                            errors.Add($"[70033] [{loc}/QIIR/Exception] TIN('{exTinVal}')이 기업구조 내 다른 CE의 TIN과 일치하지 않습니다.");
                    }
                }

                // 70034: POPE-IPE=GIR902(IPE)이고 Exception이 있으면 Art2.1.3 선택 필요
                if (ce.Qiir != null
                    && ce.Qiir.PopeIpe == Globe.PopeipeEnumType.Gir902
                    && ce.Qiir.Exception != null
                    && ce.Qiir.Exception.ExceptionRule?.Art213Specified != true)
                    errors.Add($"[70034] [{loc}/QIIR/Exception] POPE-IPE=GIR902(IPE)이고 Exception이 있으면 Art2.1.3 예외를 선택해야 합니다.");
            }

            // 70015: GIR308 존재 시, 다른 CE에 GIR307 필요
            var hasGir308 = corpStructure.Ce.Any(c => c.Id?.GlobeStatus?.Contains(Globe.IdTypeGloBeStatusEnumType.Gir308) == true);
            var hasGir307 = corpStructure.Ce.Any(c => c.Id?.GlobeStatus?.Contains(Globe.IdTypeGloBeStatusEnumType.Gir307) == true);
            if (hasGir308 && !hasGir307)
            {
                errors.Add("[70015] [1.3.2.1 구성기업] GIR308(소수지분 자회사)이 존재하지만 GIR307(소수지분모기업)이 보고된 CE가 없습니다.");
            }

            // 70019: GIR305(PE) 존재 시, GIR306(Main Entity)이 다른 CE에 필요
            var hasPe = corpStructure.Ce.Any(c => c.Id?.GlobeStatus?.Contains(Globe.IdTypeGloBeStatusEnumType.Gir305) == true);
            var hasMain = corpStructure.Ce.Any(c => c.Id?.GlobeStatus?.Contains(Globe.IdTypeGloBeStatusEnumType.Gir306) == true);
            if (hasPe && !hasMain)
            {
                errors.Add("[70019] [1.3.2.1 구성기업] GIR305(고정사업장)이 존재하지만 GIR306(본점)을 포함하는 CE가 없습니다.");
            }
        }

        #endregion

        #region Rules 일관성 (70012)

        private static void ValidateRulesConsistency(Globe.GlobeOecd globe, List<string> errors)
        {
            var corpStructure = globe.GlobeBody?.GeneralSection?.CorporateStructure;
            if (corpStructure == null) return;

            // 관할지역 → (기업명, Rules 집합) 수집
            var jurMap = new Dictionary<string, List<(string Name, HashSet<Globe.IdTypeRulesEnumType> Rules)>>();

            void AddEntry(Globe.CountryCodeType? country, string name, IList<Globe.IdTypeRulesEnumType> rules)
            {
                if (country == null || rules == null) return;
                var key = country.ToString();
                if (!jurMap.ContainsKey(key)) jurMap[key] = new List<(string, HashSet<Globe.IdTypeRulesEnumType>)>();
                jurMap[key].Add((name, new HashSet<Globe.IdTypeRulesEnumType>(rules)));
            }

            // UPE OtherUpe
            foreach (var upe in corpStructure.Upe)
            {
                if (upe.OtherUpe?.Id != null)
                {
                    var country = upe.OtherUpe.Id.ResCountryCode?.FirstOrDefault();
                    AddEntry(country, upe.OtherUpe.Id.Name ?? "(UPE)", upe.OtherUpe.Id.Rules);
                }
                if (upe.ExcludedUpe?.Id != null)
                {
                    var country = upe.ExcludedUpe.Id.ResCountryCode?.FirstOrDefault();
                    AddEntry(country, upe.ExcludedUpe.Id.Name ?? "(ExcludedUPE)", upe.ExcludedUpe.Id.Rules);
                }
            }

            // CE
            foreach (var ce in corpStructure.Ce)
            {
                if (ce.Id == null) continue;
                var country = ce.Id.ResCountryCode?.FirstOrDefault();
                AddEntry(country, ce.Id.Name ?? "(CE)", ce.Id.Rules);
            }

            // 관할지역별 Rules 일관성 검사
            foreach (var (jur, entities) in jurMap)
            {
                if (entities.Count <= 1) continue;

                // GIR204(QDMTT)를 가진 기업이 있으면 해당 관할지역 전체 면제
                bool hasQdmtt = entities.Any(e => e.Rules.Contains(Globe.IdTypeRulesEnumType.Gir204));
                if (hasQdmtt) continue;

                // Rules 집합이 모두 동일한지 확인 (GIR204 제외 후 비교)
                var first = entities[0].Rules;
                for (int i = 1; i < entities.Count; i++)
                {
                    if (!first.SetEquals(entities[i].Rules))
                    {
                        var names = string.Join(", ", entities.Select(e => e.Name));
                        errors.Add($"[70012] [1.3.x 기업구조/Rules] 관할지역 {jur} 소속 기업들({names})의 Rules가 서로 다릅니다. 동일 관할지역 기업은 동일한 Rules를 가져야 합니다.");
                        break;
                    }
                }
            }
        }

        #endregion

        #region Summary (70038, 70039, 70040)

        private static void ValidateSummary(Globe.GlobeOecd globe, List<string> errors)
        {
            var body = globe.GlobeBody;
            if (body?.Summary == null || body.Summary.Count == 0) return;

            // 보고기간 (Period End 또는 ReportingPeriod)
            var periodEnd = body.FilingInfo?.Period?.End;
            var reportingPeriod = globe.MessageSpec?.ReportingPeriod;
            var effectiveDate = (periodEnd != default ? periodEnd : reportingPeriod) ?? default;

            // UPE 국가 수집 (70040용)
            var upeCountries = new HashSet<Globe.CountryCodeType>();
            var corpStr = body.GeneralSection?.CorporateStructure;
            if (corpStr != null)
            {
                foreach (var upe in corpStr.Upe)
                {
                    if (upe.OtherUpe?.Id?.ResCountryCode != null)
                        foreach (var cc in upe.OtherUpe.Id.ResCountryCode) upeCountries.Add(cc);
                    if (upe.ExcludedUpe?.Id?.ResCountryCode != null)
                        foreach (var cc in upe.ExcludedUpe.Id.ResCountryCode) upeCountries.Add(cc);
                }
            }

            foreach (var summary in body.Summary)
            {
                var jurName = summary.Jurisdiction?.JurisdictionNameSpecified == true
                    ? summary.Jurisdiction.JurisdictionName.ToString()
                    : "(국가 없음)";

                // 70038: GIR1203/1204/1205는 2028-06-30 이후 불가
                if (effectiveDate > new DateTime(2028, 6, 30))
                {
                    foreach (var sh in summary.SafeHarbour)
                    {
                        if (sh == Globe.SafeHarbourEnumType.Gir1203
                            || sh == Globe.SafeHarbourEnumType.Gir1204
                            || sh == Globe.SafeHarbourEnumType.Gir1205)
                            errors.Add($"[70038] [1.4 요약/{jurName}] SafeHarbour {sh}은 2028.6.30 이후 사용 불가 (전환기 CbCR 적용면제 만료).");
                    }
                }

                // 70039: GIR1206은 2026-12-31 이후 불가
                if (effectiveDate > new DateTime(2026, 12, 31)
                    && summary.SafeHarbour.Contains(Globe.SafeHarbourEnumType.Gir1206))
                    errors.Add($"[70039] [1.4 요약/{jurName}] SafeHarbour GIR1206은 2026.12.31 이후 사용 불가 (UTPR 적용면제 만료).");

                // 70040: GIR1206은 UPE 국가에서만
                if (summary.SafeHarbour.Contains(Globe.SafeHarbourEnumType.Gir1206)
                    && upeCountries.Count > 0
                    && summary.Jurisdiction?.JurisdictionNameSpecified == true
                    && !upeCountries.Contains(summary.Jurisdiction.JurisdictionName))
                    errors.Add($"[70040] [1.4 요약/{jurName}] GIR1206은 UPE 국가({string.Join("/", upeCountries)})에서만 사용 가능합니다.");
            }
        }

        #endregion

        #region TIN (70001, 70002, 70003, 70005)

        private static void ValidateTin(Globe.GlobeOecd globe, List<string> errors)
        {
            // FilingCE TIN
            ValidateSingleTin(globe.GlobeBody?.FilingInfo?.FilingCe?.Tin, "1.1~1.2 신고기본정보/FilingCE", false, errors);

            // UPE TINs
            var corpStructure = globe.GlobeBody?.GeneralSection?.CorporateStructure;
            if (corpStructure == null) return;

            foreach (var upe in corpStructure.Upe)
            {
                if (upe.OtherUpe?.Id != null)
                {
                    var upeName = upe.OtherUpe.Id.Name ?? "(이름없음)";
                    if (upe.OtherUpe.Id.Tin.Count == 0)
                        errors.Add($"[70005] [1.3.1 최종모기업/OtherUPE '{upeName}'] TIN이 없습니다. UPE는 실제 TIN이 필수입니다.");
                    else
                        foreach (var tin in upe.OtherUpe.Id.Tin)
                            ValidateSingleTin(tin, $"1.3.1 최종모기업/OtherUPE '{upeName}'", true, errors);
                }
                if (upe.ExcludedUpe?.Id != null)
                {
                    var upeName = upe.ExcludedUpe.Id.Name ?? "(이름없음)";
                    if (upe.ExcludedUpe.Id.Tin.Count == 0)
                        errors.Add($"[70005] [1.3.1 최종모기업/ExcludedUPE '{upeName}'] TIN이 없습니다.");
                    else
                        foreach (var tin in upe.ExcludedUpe.Id.Tin)
                            ValidateSingleTin(tin, $"1.3.1 최종모기업/ExcludedUPE '{upeName}'", false, errors);
                }
            }

            // CE TINs
            foreach (var ce in corpStructure.Ce)
            {
                if (ce.Id?.Tin != null)
                {
                    var ceName = ce.Id?.Name ?? "(이름없음)";
                    foreach (var tin in ce.Id.Tin)
                        ValidateSingleTin(tin, $"1.3.2.1 구성기업/CE '{ceName}'", false, errors);

                    // 70005: GIR316/GIR318이 아닌 CE TIN에도 NOTIN 불가
                    bool isNonGroupOrExited = ce.Id?.GlobeStatus?.Any(s =>
                        s == Globe.IdTypeGloBeStatusEnumType.Gir316
                        || s == Globe.IdTypeGloBeStatusEnumType.Gir318) == true;
                    if (!isNonGroupOrExited)
                    {
                        foreach (var tin in ce.Id.Tin)
                        {
                            if (string.Equals(tin.Value, "NOTIN", StringComparison.OrdinalIgnoreCase)
                                || (tin.TypeOfTinSpecified && tin.TypeOfTin == Globe.TinEnumType.Gir3004)
                                || (tin.UnknownSpecified && tin.Unknown))
                                errors.Add($"[70005] [1.3.2.1 구성기업/CE '{ceName}'] GIR316/GIR318 상태가 아닌 CE TIN에는 NOTIN/GIR3004/Unknown=TRUE가 허용되지 않습니다.");
                        }
                    }
                }
            }
        }

        private static void ValidateSingleTin(Globe.TinType tin, string context, bool isUpe, List<string> errors)
        {
            if (tin == null) return;

            var isNoTin = string.Equals(tin.Value, "NOTIN", StringComparison.OrdinalIgnoreCase);
            var isGir3004 = tin.TypeOfTinSpecified && tin.TypeOfTin == Globe.TinEnumType.Gir3004;
            var isUnknown = tin.UnknownSpecified && tin.Unknown;

            // 70001: GIR3004이면 TIN='NOTIN', Unknown=TRUE, IssuedBy 없음
            if (isGir3004 && (!isNoTin || !isUnknown))
            {
                errors.Add($"[70001] [{context}] TypeOfTIN=GIR3004이면 TIN='NOTIN', Unknown=TRUE이어야 합니다.");
            }
            if (isGir3004 && tin.IssuedBySpecified)
            {
                errors.Add($"[70001] [{context}] TypeOfTIN=GIR3004이면 IssuedBy를 제공하면 안 됩니다.");
            }

            // 70002: TIN='NOTIN'이면 GIR3004, Unknown=TRUE, IssuedBy 미제공
            if (isNoTin && (!isGir3004 || !isUnknown || tin.IssuedBySpecified))
            {
                errors.Add($"[70002] [{context}] TIN='NOTIN'이면 TypeOfTIN=GIR3004, Unknown=TRUE이어야 하고 IssuedBy는 생략해야 합니다.");
            }

            // 70003: Unknown=TRUE이면 NOTIN, GIR3004, IssuedBy 미제공
            if (isUnknown && (!isNoTin || !isGir3004 || tin.IssuedBySpecified))
            {
                errors.Add($"[70003] [{context}] Unknown=TRUE이면 TIN='NOTIN', TypeOfTIN=GIR3004이어야 하고 IssuedBy는 생략해야 합니다.");
            }

            // 70005: UPE TIN에는 GIR3004/Unknown=TRUE 불가
            if (isUpe && (isGir3004 || isUnknown))
            {
                errors.Add($"[70005] [{context}] UPE TIN에는 TypeOfTIN=GIR3004 또는 Unknown=TRUE가 허용되지 않습니다.");
            }

            // 70007: TypeOfTIN=GIR3003이면 P2JJYYYYMMDDCCCXXX 형식
            if (tin.TypeOfTinSpecified && tin.TypeOfTin == Globe.TinEnumType.Gir3003
                && !string.IsNullOrEmpty(tin.Value))
            {
                // P2 + 2자리국가코드 + 8자리날짜 + 3자리그룹코드 + 3자리고유번호 = 18자
                if (!Regex.IsMatch(tin.Value, @"^P2[A-Z]{2}\d{8}[A-Z0-9]{3}[A-Z0-9]{3}$"))
                {
                    errors.Add($"[70007] [{context}] TypeOfTIN=GIR3003의 TIN 형식이 올바르지 않습니다. 형식: P2[국가코드2자][날짜8자][그룹코드3자][고유번호3자] (현재: '{tin.Value}')");
                }
            }
        }

        #endregion
    }
}
