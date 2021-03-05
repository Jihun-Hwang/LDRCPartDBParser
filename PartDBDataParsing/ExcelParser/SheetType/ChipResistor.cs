using Pentacube.CAx.ECAD.LcadIf;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PartDBDataParsing.ExcelParser.SheetType
{
    /// <summary>
    /// Classification : Resistor
    /// </summary>
    public sealed partial class ChipResistor : CommonExcelData
    {
        // row index
        private const int PackageSizeIdx = 4;
        private const int ResistanceValueUnitIdx = 16;
        private const int PowerIdx = 17;
        private const int ToleranceIdx = 18;

        // 고유 속성 (셀에 있으면 0, 없으면 x)
        public double ResistanceValue { get; set; }                        // 저항값 - o
        public eResistanceUnit ResistanceUnit { get; set; } = default;     // 저항단위 - o
        public double Tolerance { get; set; } = default;                   // 허용 오차 - o
        public string PackageSize { get; set; } = string.Empty;            // 부품 크기 - o
        public bool ArrayType { get; set; } = default;                     // ArrayType - x
        public double Power { get; set; } = default;                       // 정격 전력 - o

        public ChipResistor() : base(ePartClassification.Resistor)
        {
        }

        public override XElement CreateXmlElement()
        {
            var attributes = new[]
            {
                new XAttribute(nameof(ResistanceValue), ResistanceValue),
                new XAttribute(nameof(ResistanceUnit), ResistanceUnit),
                new XAttribute(nameof(Tolerance), Tolerance),
                new XAttribute(nameof(PackageSize), PackageSize),
                new XAttribute(nameof(ArrayType), ArrayType),
                new XAttribute(nameof(Power), Power),
            };

            return CreateXmlElement(attributes);
        }
    }

    public partial class ChipResistor
    {
        public static IEnumerable<ChipResistor> Parse(ParsedExcelData excelData)
        {
            if (excelData.Context.Count == 0)
                yield break;

            foreach (var row in excelData.Context.Skip(1))
                yield return GetChipResistor(row); ;
        }

        private static ChipResistor GetChipResistor(string[] row)
        {
            var resistanceePair = ConvertUnit(row[ResistanceValueUnitIdx]);

            return new ChipResistor
            {
                // 공통 property
                VendorCode = row[VendorCodeIdx],
                CustomerCode = row[CustomerCodeIdx],

                // 고유 속성 삽입
                //PackageSize = row[PackageSizeIdx],     // TODO: chip size col 보류
                ResistanceValue = resistanceePair.Value,
                ResistanceUnit = ConvertSItoResistanceUnit(resistanceePair.Prefix),
                Power = ParseToDoubleOrDefault(row[PowerIdx]),
                Tolerance = ParseToDoubleOrDefault(row[ToleranceIdx], true),
            };
        }
    }
}
