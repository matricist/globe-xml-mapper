using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace GlobeMapper.Services
{
    public static class XmlExportService
    {
        private static readonly XmlSerializerNamespaces Namespaces;
        private static readonly XmlSerializer Serializer;

        static XmlExportService()
        {
            Serializer = new XmlSerializer(typeof(Globe.GlobeOecd));

            Namespaces = new XmlSerializerNamespaces();
            Namespaces.Add("globe", "urn:oecd:ties:globe:v2");
            Namespaces.Add("iso", "urn:oecd:ties:isoglobetypes:v1");
            Namespaces.Add("oecd", "urn:oecd:ties:oecdglobetypes:v1");
            Namespaces.Add("stf", "urn:oecd:ties:globestf:v5");
        }

        /// <summary>
        /// GlobeOecd 객체를 XML 문자열로 직렬화 (UTF-8 선언).
        /// </summary>
        public static string Serialize(Globe.GlobeOecd globe)
        {
            var settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "  ",
                Encoding = new UTF8Encoding(false), // BOM 제외
                OmitXmlDeclaration = false,
            };

            // StringWriter 는 내부 UTF-16 이라 XML 선언이 "utf-16" 으로 새어 나옴 →
            // MemoryStream + UTF-8 XmlWriter 경유하여 실제 인코딩과 선언 일치시킴.
            using var ms = new MemoryStream();
            using (var xw = XmlWriter.Create(ms, settings))
            {
                Serializer.Serialize(xw, globe, Namespaces);
            }
            return Encoding.UTF8.GetString(ms.ToArray());
        }

        /// <summary>
        /// GlobeOecd 객체를 XML 파일로 저장
        /// </summary>
        public static void SerializeToFile(Globe.GlobeOecd globe, string outputPath)
        {
            var xml = Serialize(globe);
            File.WriteAllText(outputPath, xml, new UTF8Encoding(false));
        }
    }
}
