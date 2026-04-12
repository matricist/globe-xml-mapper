/// 샘플 폴더에서 XML 변환 테스트
using GlobeMapper.Services;
using System.Xml.Serialization;

var sampleRoot = args.Length > 0 ? args[0]
    : Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "sample");

var globe = new Globe.GlobeOecd
{
    Version = "2.0",
    MessageSpec = new Globe.MessageSpecType(),
    GlobeBody = new Globe.GlobeBodyType()
};
var orch = new MappingOrchestrator();

Console.WriteLine($"=== MapFolder: {sampleRoot} ===");
var errors = orch.MapFolder(sampleRoot, globe);

Console.WriteLine($"\n오류 {errors.Count}건:");
foreach (var e in errors) Console.WriteLine($"  {e}");

Console.WriteLine("\n=== JurisdictionSection ===");
foreach (var js in globe.GlobeBody.JurisdictionSection)
{
    Console.WriteLine($"  Jurisdiction={js.Jurisdiction}  JurWithTaxingRights={js.JurWithTaxingRights.Count}건");
    foreach (var jw in js.JurWithTaxingRights)
    {
        Console.WriteLine($"    -> {jw.JurisdictionName}  ReportDifference={(jw.ReportDifference != null ? "있음" : "없음")}");
        if (jw.ReportDifference is { } rd)
        {
            if (rd.EtrDifferenceSpecified) Console.WriteLine($"       ETRDifference={rd.EtrDifference}");
            if (rd.TuTDifference != null)  Console.WriteLine($"       TuTDifference={rd.TuTDifference}");
            if (rd.ElectionsDifference != null) Console.WriteLine($"       Elections={rd.ElectionsDifference}");
            if (rd.TransitionDifferenceSpecified) Console.WriteLine($"       TransitionDifference={rd.TransitionDifference}");
        }
    }
    foreach (var etr in js.GLoBeTax.Etr)
    {
        var oc = etr.EtrStatus?.EtrComputation?.OverallComputation;
        if (oc != null)
        {
            Console.WriteLine($"  ETR: FANIL={oc.Fanil} NetGlobeIncome.Total={oc.NetGlobeIncome?.Total}");
            Console.WriteLine($"       TopUpTax={oc.TopUpTax} ExcessProfits={oc.ExcessProfits}");
            if (oc.Qdmtt != null) Console.WriteLine($"       QDMTT: Amount={oc.Qdmtt.Amount} MinRate={oc.Qdmtt.MinRate}");
            if (oc.AdditionalTopUpTax?.NonArt415?.Count > 0)
                Console.WriteLine($"       AdditionalTopUpTax.NonArt415: {oc.AdditionalTopUpTax.NonArt415.Count}건");
            if (oc.SubstanceExclusion != null)
                Console.WriteLine($"       SubstanceExclusion.Total={oc.SubstanceExclusion.Total}");
            var ce = etr.EtrStatus?.EtrComputation?.CeComputation?.FirstOrDefault();
            if (ce?.NetGlobeIncome?.IntShippingIncome != null)
                Console.WriteLine($"       IntShippingIncome 있음");
        }
    }
}

// XML 직렬화
var ns = new XmlSerializerNamespaces();
ns.Add("globe", "urn:oecd:ties:globe:v2");
ns.Add("stf",   "urn:oecd:ties:globestf:v5");
ns.Add("oecd",  "urn:oecd:ties:oecdglobetypes:v1");
ns.Add("iso",   "urn:oecd:ties:isoglobetypes:v1");

var outPath = Path.Combine(sampleRoot, "TEST_OUTPUT.xml");
using var fs = new FileStream(outPath, FileMode.Create);
new XmlSerializer(typeof(Globe.GlobeOecd)).Serialize(fs, globe, ns);
Console.WriteLine($"\nXML 저장: {outPath}");
