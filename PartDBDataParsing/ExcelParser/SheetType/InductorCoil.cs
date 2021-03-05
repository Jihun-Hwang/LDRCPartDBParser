using Pentacube.CAx.ECAD.LcadIf;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PartDBDataParsing.ExcelParser.SheetType
{
    /// <summary>
    /// Classification : Inductor
    /// </summary>
    public sealed partial class InductorCoil : CommonExcelData
    {
        // row index
        private const int InductanceValueUnitIdx = 9;
        private const int MaxCurrentIdx = 13;

        // 고유 속성
        public double InductanceValue { get; set; }                      // 유도 용량 값 - o
        public eInductorUnit InductanceUnit { get; set; } = default;     // 유도 용량 단위 - o
        public bool IsCmChoke { get; set; } = default;                   // CM Choke 여부 - x
        public double MaxCurrent { get; set; }                           // 최대허용전류 - o
        public double CapacitanceValue { get; set; }                     // 정전 용량 값 - o
        public eCapacitanceUnit CapacitanceUnit { get; set; } = default; // 정전 용량 단위 - o

        public InductorCoil() : base(ePartClassification.Inductor)
        {
        }

        public override XElement CreateXmlElement()
        {
            var attributes = new[]
            {
                new XAttribute(nameof(InductanceValue), InductanceValue),
                new XAttribute(nameof(InductanceUnit), InductanceUnit),
                new XAttribute(nameof(IsCmChoke), IsCmChoke),
                new XAttribute(nameof(MaxCurrent), MaxCurrent),
                new XAttribute(nameof(CapacitanceValue), CapacitanceValue),
                new XAttribute(nameof(CapacitanceUnit), CapacitanceUnit),
            };

            return CreateXmlElement(attributes);
        }
    }

    public partial class InductorCoil
    {
        public static IEnumerable<InductorCoil> Parse(ParsedExcelData excelData)
        {
            if (excelData.Context.Count == 0)
                yield break;

            foreach (var row in excelData.Context.Skip(1))
                yield return GetInductorCoil(row);
        }

        private static InductorCoil GetInductorCoil(string[] row)
        {
            //TODO : inductance또는 capacitance 중 하나는 잘못 매핑된 것
            var inductancePair = ConvertUnit(row[InductanceValueUnitIdx]);
            var capacitancePair = ConvertUnit(row[InductanceValueUnitIdx]);

            return new InductorCoil
            {
                // 공통 property
                VendorCode = row[VendorCodeIdx],
                CustomerCode = row[CustomerCodeIdx],

                // 고유 속성 삽입
                InductanceValue = inductancePair.Value,
                InductanceUnit = ConvertSItoInductorUnit(inductancePair.Prefix),
                CapacitanceValue = capacitancePair.Value,
                CapacitanceUnit = ConvertSItoCapacitanceUnit(capacitancePair.Prefix),
                MaxCurrent = ParseToDoubleOrDefault(row[MaxCurrentIdx]),
            };
        }
    }
}
