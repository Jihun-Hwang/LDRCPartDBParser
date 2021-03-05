using Pentacube.CAx.ECAD.LcadIf;
using System.Xml.Linq;

namespace PartDBDataParsing.ExcelParser.SheetType.PinType
{
    public class ICPin : IPin
    {
        public string Num { get; set; } = string.Empty;
        public eICPinType Type { get; set; } = default;
        public string Group { get; set; } = string.Empty;
        /// <summary>
        ///  전압
        /// </summary>
        public double Voltage { get; set; }
        /// <summary>
        /// 전류
        /// </summary>
        public double Current { get; set; }
        /// <summary>
        /// 최소작동전압
        /// </summary>
        public double VoltageMin { get; set; }
        /// <summary>
        /// 최대작동전압
        /// </summary>
        public double VoltageMax { get; set; }
        /// <summary>
        /// 한계전압
        /// </summary>
        public double VoltageLimit { get; set; }
        /// <summary>
        /// Wetting Current
        /// </summary>
        public double WettingCurrent { get; set; }
        /// <summary>
        /// Esd 보호 여부
        /// </summary>
        public bool EsdProtection { get; set; }
        /// <summary>
        /// Short 보호 여부
        /// </summary>
        public bool ShortProtection { get; set; }
        /// <summary>
        /// 환류다이오드
        /// </summary>
        public bool FreeWheelingProtection { get; set; }

        public XElement CreatePinXmlElement()
        {
            return new XElement(
                "Pin",
                new XAttribute(nameof(Num), Num),
                new XAttribute(nameof(Type), Type),
                new XAttribute(nameof(Group), Group),
                new XAttribute(nameof(Voltage), Voltage),
                new XAttribute(nameof(Current), Current),
                new XAttribute(nameof(VoltageMin), VoltageMin),
                new XAttribute(nameof(VoltageMax), VoltageMax),
                new XAttribute(nameof(VoltageLimit), VoltageLimit),
                new XAttribute(nameof(WettingCurrent), WettingCurrent),
                new XAttribute(nameof(EsdProtection), EsdProtection),
                new XAttribute(nameof(ShortProtection), ShortProtection),
                new XAttribute(nameof(FreeWheelingProtection), FreeWheelingProtection));
        }
    }
}
