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
        /// GlobeOecd 객체를 XML 문자열로 직렬화
        /// </summary>
        public static string Serialize(Globe.GlobeOecd globe)
        {
            var settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "  ",
                Encoding = Encoding.UTF8,
                OmitXmlDeclaration = false
            };

            using var sw = new StringWriter();
            using var xw = XmlWriter.Create(sw, settings);
            Serializer.Serialize(xw, globe, Namespaces);
            return sw.ToString();
        }

        /// <summary>
        /// GlobeOecd 객체를 XML 파일로 저장
        /// </summary>
        public static void SerializeToFile(Globe.GlobeOecd globe, string outputPath)
        {
            var xml = Serialize(globe);
            File.WriteAllText(outputPath, xml, Encoding.UTF8);
        }
    }
}
