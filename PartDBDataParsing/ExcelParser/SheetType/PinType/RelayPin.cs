using Pentacube.CAx.ECAD.LcadIf;
using System.Xml.Linq;

namespace PartDBDataParsing.ExcelParser.SheetType.PinType
{
    public class RelayPin : IPin
    {
        public string Num { get; set; } = string.Empty;
        public eRelayPin Type { get; set; } = default;

        public XElement CreatePinXmlElement()
        {
            return new XElement(
                "Pin",
                new XAttribute(nameof(Num), Num),
                new XAttribute(nameof(Type), Type));
        }
    }
}
