using Pentacube.CAx.ECAD.LcadIf;
using System.Xml.Linq;

namespace PartDBDataParsing.ExcelParser.SheetType.PinType
{
    public class FETPin : IPin
    {
        public string Num { get; set; } = string.Empty;
        public eFETPinType Type { get; set; } = default;
        public string Group { get; set; } = string.Empty;

        public XElement CreatePinXmlElement()
        {
            return new XElement(
                "Pin",
                new XAttribute(nameof(Num), Num),
                new XAttribute(nameof(Type), Type),
                new XAttribute(nameof(Group), Group));
        }
    }
}
