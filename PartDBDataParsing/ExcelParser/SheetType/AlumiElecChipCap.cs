using Pentacube.CAx.ECAD.LcadIf;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PartDBDataParsing.ExcelParser.SheetType
{
    /// <summary>
    /// Classification : Electrolytic Capacitor(AlCapacitor)
    /// </summary>
    public sealed partial class AlumiElecChipCap : CommonExcelData
    {
        // row index
        private const int VoltageIdx = 5;
        private const int CapacitanceValueUnitIdx = 11;
        private const int ToleranceIdx = 13;

        // 고유 속성
        public double CapacitanceValue { get; set; }              // 정전 용량 값 - o
        public eCapacitanceUnit CapacitanceUnit { get; set; }     // 정전 용량 단위 - o
        public double Tolerance { get; set; }                     // 허용오차 - o
        public bool OpenType { get; set; }                        // OpenType - x
        public bool ArrayType { get; set; }                       // ArrayType - x
        public double EnduranceTemp { get; set; }                 // 수명온도 - x
        public double EnduranceTime { get; set; }                 // 수명시간 - x
        public double PackageDiameter { get; set; }               // 부품 지름 - x
        public double PackageHeight { get; set; }                 // 부품 높이(mm) - x
        public string AnodePinNum { get; set; } = string.Empty;   // +극 - x
        public double Voltage { get; set; }                       // 정격전압 - o

        public AlumiElecChipCap() : base(ePartClassification.AlCapacitor)
        {
        }

        public override XElement CreateXmlElement()
        {
            var attributes = new[]
            {
                new XAttribute(nameof(CapacitanceValue), CapacitanceValue),
                new XAttribute(nameof(CapacitanceUnit), CapacitanceUnit),
                new XAttribute(nameof(Tolerance), Tolerance),
                new XAttribute(nameof(OpenType), OpenType),
                new XAttribute(nameof(ArrayType), ArrayType),
                new XAttribute(nameof(EnduranceTemp), EnduranceTemp),
                new XAttribute(nameof(EnduranceTime), EnduranceTime),
                new XAttribute(nameof(PackageDiameter), PackageDiameter),
                new XAttribute(nameof(PackageHeight), PackageHeight),
                new XAttribute(nameof(AnodePinNum), AnodePinNum),
                new XAttribute(nameof(Voltage), Voltage)
            };

            return CreateXmlElement(attributes);
        }
    }

    public partial class AlumiElecChipCap
    {
        public static IEnumerable<AlumiElecChipCap> Parse(ParsedExcelData excelData)
        {
            if (excelData.Context.Count == 0)
                yield break;

            foreach (var row in excelData.Context.Skip(1))
                yield return GetAlumiElecChipCap(row);
        }

        private static AlumiElecChipCap GetAlumiElecChipCap(string[] row)
        {
            var capacitancePair = ConvertUnit(row[CapacitanceValueUnitIdx]);

            return new AlumiElecChipCap
            {
                // 공통 property
                VendorCode = row[VendorCodeIdx],
                CustomerCode = row[CustomerCodeIdx],

                // 고유 속성 삽입
                Voltage = ConvertUnit(row[VoltageIdx]).Value,
                CapacitanceValue = capacitancePair.Value,
                CapacitanceUnit = ConvertSItoCapacitanceUnit(capacitancePair.Prefix),
                Tolerance = ParseToDoubleOrDefault(row[ToleranceIdx], true),
            };
        }

    }
}
