using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using ClosedXML.Excel;

namespace GlobeMapper.Services
{
    public abstract class MappingBase
    {
        protected SheetMapping Mapping { get; }

        protected MappingBase(string mappingFileName)
        {
            var jsonPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory, "Resources", "mappings", mappingFileName);
            var json = File.ReadAllText(jsonPath);
            Mapping = JsonSerializer.Deserialize<SheetMapping>(
                json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        }

        public string SheetName => Mapping.SheetName;
        public bool Repeatable => Mapping.Repeatable;

        public abstract void Map(IXLWorksheet ws, Globe.GlobeOecd globe, List<string> errors, string fileName);

        #region 공통 유틸리티

        protected static void ForEachValue(IXLWorksheet ws, MappingEntry m, string fileName,
            List<string> errors, Action<string> action)
        {
            try
            {
                var cellValue = ws.Cell(m.Cell).GetString()?.Trim();
                if (string.IsNullOrEmpty(cellValue)) return;

                var values = m.Multi
                    ? cellValue.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    : new[] { cellValue };

                foreach (var val in values)
                    action(val);
            }
            catch (Exception ex)
            {
                errors.Add($"[{fileName}] 셀 {m.Cell} ({m.Label}) 매핑 오류: {ex.Message}");
            }
        }

        protected static void SetEnum<T>(string value, Action<T> setter,
            List<string> errors, string fileName, MappingEntry entry) where T : struct, Enum
        {
            if (TryParseEnum<T>(value, out var result))
                setter(result);
            else
                errors.Add($"[{fileName}] 셀 {entry.Cell}: {typeof(T).Name} 변환 실패 '{value}'");
        }

        protected static bool TryParseDate(string value, out DateTime result)
            => DateTime.TryParse(value, out result);

        protected static bool ParseBool(string value)
        {
            if (bool.TryParse(value, out var b)) return b;
            var v = value.Trim().ToUpper();
            return v == "Y" || v == "YES" || v == "1" || v == "TRUE" || v == "O" || v == "예";
        }

        protected static bool TryParseEnum<T>(string value, out T result) where T : struct, Enum
        {
            if (Enum.TryParse(value, true, out result))
                return true;

            foreach (var field in typeof(T).GetFields(
                System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static))
            {
                var xmlAttr = field.GetCustomAttributes(
                    typeof(System.Xml.Serialization.XmlEnumAttribute), false);
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
