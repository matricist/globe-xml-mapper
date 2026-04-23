using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    public abstract class MappingBase
    {
        protected SheetMapping Mapping { get; }

        protected MappingBase(string mappingFileName)
        {
            if (mappingFileName == null)
                return;
            var jsonPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Resources",
                "mappings",
                mappingFileName
            );
            var json = File.ReadAllText(jsonPath);
            Mapping = JsonSerializer.Deserialize<SheetMapping>(
                json,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true }
            );
        }

        public string SheetName => Mapping?.SheetName;
        public bool Repeatable => Mapping?.Repeatable ?? false;

        public abstract void Map(
            IXLWorksheet ws,
            Globe.GlobeOecd globe,
            List<string> errors,
            string fileName
        );

        #region 공통 유틸리티

        protected static void ForEachValue(
            IXLWorksheet ws,
            MappingEntry m,
            string fileName,
            List<string> errors,
            Action<string> action
        )
        {
            try
            {
                var cellValue = ws.Cell(m.Cell).GetString()?.Trim();
                if (string.IsNullOrEmpty(cellValue))
                    return;

                var values = m.Multi
                    ? cellValue.Split(
                        ',',
                        StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries
                    )
                    : new[] { cellValue };

                foreach (var val in values)
                    action(val);
            }
            catch (Exception ex)
            {
                errors.Add($"[{fileName}] 셀 {m.Cell} ({m.Label}) 매핑 오류: {ex.Message}");
            }
        }

        protected static void SetEnum<T>(
            string value,
            Action<T> setter,
            List<string> errors,
            string fileName,
            MappingEntry entry
        )
            where T : struct, Enum
        {
            if (TryParseEnum<T>(value, out var result))
                setter(result);
            else
            {
                // 진단: 실제 바이트 값도 포함 (숨은 문자 확인용)
                var bytes = string.Join(
                    " ",
                    System.Text.Encoding.UTF8.GetBytes(value).Select(b => b.ToString("X2"))
                );
                errors.Add(
                    $"[{fileName}] 셀 {entry.Cell}: {typeof(T).Name} 변환 실패 '{value}' (bytes: {bytes})"
                );
            }
        }

        protected static bool TryParseDate(string value, out DateTime result) =>
            DateTime.TryParse(value, out result);

        /// <summary>
        /// TIN 미보유 시 사용. &lt;TIN TypeOfTIN="GIR3004" unknown="true"&gt;NOTIN&lt;/TIN&gt;
        /// UPE(OtherUPE)에는 사용 불가 (에러 70005) — CE 등 non-UPE 전용.
        /// </summary>
        protected static Globe.TinType NoTin() =>
            new Globe.TinType
            {
                Value = "NOTIN",
                Unknown = true,
                UnknownSpecified = true,
                TypeOfTin = Globe.TinEnumType.Gir3004,
                TypeOfTinSpecified = true,
            };

        /// <summary>
        /// "번호,유형코드,발급국가" 형식 파싱. 유형·발급국가는 생략 가능.
        /// 예: "1234567890,GIR3001,KR"
        /// 빈 값이나 "NOTIN" 입력 시 unknown="true" 처리.
        /// </summary>
        protected static Globe.TinType ParseTin(string input)
        {
            if (
                string.IsNullOrWhiteSpace(input)
                || input.Trim().Equals("NOTIN", StringComparison.OrdinalIgnoreCase)
            )
                return NoTin();

            var parts = input.Split(',', StringSplitOptions.TrimEntries);
            var tin = new Globe.TinType { Value = parts[0] };

            if (
                parts.Length >= 2
                && !string.IsNullOrEmpty(parts[1])
                && TryParseEnum<Globe.TinEnumType>(parts[1], out var tinType)
            )
            {
                tin.TypeOfTin = tinType;
                tin.TypeOfTinSpecified = true;
            }

            if (
                parts.Length >= 3
                && !string.IsNullOrEmpty(parts[2])
                && TryParseEnum<Globe.CountryCodeType>(parts[2], out var country)
            )
            {
                tin.IssuedBy = country;
                tin.IssuedBySpecified = true;
            }

            return tin;
        }

        /// <summary>
        /// "영문;국문" 형식 파싱. ';' 없으면 (value, null).
        /// </summary>
        protected static (string Name, string KName) ParseNameKName(string value)
        {
            var parts = value.Split(';', 2, StringSplitOptions.TrimEntries);
            return parts.Length == 2 ? (parts[0], parts[1]) : (parts[0], null);
        }

        protected static bool ParseBool(string value)
        {
            if (bool.TryParse(value, out var b))
                return b;
            var v = value.Trim().ToUpper();
            return v == "Y"
                || v == "YES"
                || v == "1"
                || v == "TRUE"
                || v == "O"
                || v == "예"
                || v == "여";
        }

        protected static bool TryParseEnum<T>(string value, out T result)
            where T : struct, Enum
        {
            if (TryParseEnumCore<T>(value, out result))
                return true;

            // "GIR701 구성기업" 같이 코드 뒤에 설명이 붙은 경우 첫 단어만 시도
            var firstWord = value.Split(' ', 2)[0].Trim();
            if (
                firstWord.Length > 0
                && firstWord != value
                && TryParseEnumCore<T>(firstWord, out result)
            )
                return true;

            // 비가시 유니코드 문자(NBSP, 제로폭 공백 등) 제거 후 재시도
            var cleaned = new string(
                value
                    .Where(c =>
                        !char.IsControl(c) && c != '\u00A0' && c != '\u200B' && c != '\uFEFF'
                    )
                    .ToArray()
            ).Trim();
            if (cleaned.Length > 0 && cleaned != value && TryParseEnumCore<T>(cleaned, out result))
                return true;

            result = default;
            return false;
        }

        private static bool TryParseEnumCore<T>(string value, out T result)
            where T : struct, Enum
        {
            // XmlEnum 매핑 먼저 시도 — OECD GIR 코드("GIR701" 등)는 XmlEnumAttribute로 정의됨
            // Enum.TryParse의 케이스 인센시티브 동작에 의존하지 않음
            foreach (
                var field in typeof(T).GetFields(
                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static
                )
            )
            {
                var xmlAttr = field.GetCustomAttributes(
                    typeof(System.Xml.Serialization.XmlEnumAttribute),
                    false
                );
                if (xmlAttr.Length > 0)
                {
                    var xmlValue = ((System.Xml.Serialization.XmlEnumAttribute)xmlAttr[0]).Name;
                    if (string.Equals(xmlValue, value, StringComparison.OrdinalIgnoreCase))
                    {
                        result = (T)field.GetValue(null);
                        return true;
                    }
                }
            }

            // XmlEnum에 없으면 enum 멤버 이름으로 시도
            if (Enum.TryParse(value, true, out result))
                return true;

            result = default;
            return false;
        }

        #endregion
    }

    #region JSON 모델

    public class SheetMapping
    {
        public string Description { get; set; }
        public string SheetName { get; set; }
        public bool Repeatable { get; set; }
        public string CollectionTarget { get; set; }
        public Dictionary<string, SectionMapping> Sections { get; set; }
    }

    public class SectionMapping
    {
        public string Description { get; set; }
        public List<MappingEntry> Mappings { get; set; }
    }

    public class MappingEntry
    {
        public string Cell { get; set; }
        public string Target { get; set; }
        public string Type { get; set; }
        public string Label { get; set; }
        public bool Multi { get; set; }
    }

    #endregion
}
