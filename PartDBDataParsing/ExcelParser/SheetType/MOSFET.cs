using PartDBDataParsing.ExcelParser.SheetType.PinType;
using Pentacube.CAx.ECAD.LcadIf;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PartDBDataParsing.ExcelParser.SheetType
{
    /// <summary>
    /// Classification : FET
    /// </summary>
    public sealed partial class MOSFET : CommonExcelData, ICommonPin
    {
        // row index
        private const int VdsIdx = 15;
        private const int VgsIdx = 17;

        // 고유 속성
        public eFETType Type { get; set; } = default;             // 트랜지스터 유형 - x
        public double DrainCurrent { get; set; }                  // Drain 전류(A) - x
        public double Vgs { get; set; }                           // 최대 Vgs 전압 - o
        public double Vds { get; set; }                           // 최대 Vds 전압 - o
        public bool ArrayType { get; set; } = default;            // ArrayType - x
        public double Power { get; set; }                         // 정격 전력 - x

        // PIN
        public List<IPin> Pins { get; } = new List<IPin>();

        public MOSFET() : base(ePartClassification.Fet)
        {
        }

        public override XElement CreateXmlElement()
        {
            var attributes = new[]
            {
                new XAttribute(nameof(Type), Type),
                new XAttribute(nameof(DrainCurrent), DrainCurrent),
                new XAttribute(nameof(Vgs), Vgs),
                new XAttribute(nameof(Vds), Vds),
                new XAttribute(nameof(ArrayType), ArrayType),
                new XAttribute(nameof(Power), Power),
            };

            return CreateXmlElement(attributes);
        }
    }

    public partial class MOSFET
    {
        public static IEnumerable<MOSFET> Parse(ParsedExcelData excelData)
        {
            if (excelData.Context.Count == 0)
                yield break;

            foreach (var row in excelData.Context.Skip(1))
                yield return GetMOSFET(row);
        }

        private static MOSFET GetMOSFET(string[] row)
        {
            var vdsIdxPair = ConvertUnit(row[VdsIdx]);
            var vgsIdxPair = ConvertUnit(row[VgsIdx]);

            return new MOSFET
            {
                // 공통 property
                VendorCode = row[VendorCodeIdx],
                CustomerCode = row[CustomerCodeIdx],

                // 고유 속성 삽입
                Vds = vdsIdxPair.Value,
                Vgs = vgsIdxPair.Value,
            };
        }
    }
}
