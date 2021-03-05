using System.Xml.Linq;

namespace PartDBDataParsing.ExcelParser.SheetType.PinType
{
    public interface IPinSerialize
    {
        XElement CreatePinXmlElement();
    }

    public interface IPin : IPinSerialize
    {
        string Num { get; }
    }
}
