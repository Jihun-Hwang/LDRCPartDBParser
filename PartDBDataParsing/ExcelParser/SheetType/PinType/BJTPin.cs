using Pentacube.CAx.ECAD.LcadIf;
using System.Xml.Linq;

namespace PartDBDataParsing.ExcelParser.SheetType.PinType
{
    public class BJTPin : IPin
    {
        public string Num { get; set; } = string.Empty;
        public eBJTPinType Type { get; set; } = default; // BASE COLLECTOR EMITTER
        public string Group { get; set; } = string.Empty;

        public BJTPin() : this(string.Empty, eBJTPinType.BASE) { }
        public BJTPin((string, eBJTPinType)value) : this(value.Item1, value.Item2) { }
        public BJTPin(string num, eBJTPinType type)
        {
            Num = num;
            Type = type;
        }

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
