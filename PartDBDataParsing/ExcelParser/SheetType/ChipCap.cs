using Pentacube.CAx.ECAD.LcadIf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PartDBDataParsing.ExcelParser.SheetType
{
    /// <summary>
    /// Classification : MLCC
    /// </summary>
    public sealed partial class ChipCap : CommonExcelData
    {
        // row index
        private const int PackageSizeIdx = 4;
        private const int VoltageIdx = 5;
        private const int CapacitanceValueUnitIdx = 14;
        private const int ToleranceIdx = 16;

        // 고유 속성
        public double CapacitanceValue { get; set; }                       // 정전 용량 값 - o
        public eCapacitanceUnit CapacitanceUnit { get; set; } = default;   // 정전 용량 단위 - o
        public double Tolerance { get; set; }                              // 허용 오차 - o
        public double PackageSize { get; set; }                            // 부품 크기 - o
        public bool OpenType { get; set; } = default;                      // OpenType - X
        public bool ArrayType { get; set; } = default;                     // ArrayType - X
        public eTemperatureCharacter Temperature { get; set; } = default;  // 온도 단위 default = CH - X
        public double Voltage { get; set; }                                // 정격 전압 - o

        public ChipCap() : base(ePartClassification.MlccCapacitor)
        {
        }

        public override XElement CreateXmlElement()
        {
            var attributes = new[]
            {
                new XAttribute(nameof(CapacitanceValue), CapacitanceValue),
                new XAttribute(nameof(CapacitanceUnit), CapacitanceUnit),
                new XAttribute(nameof(Tolerance), Tolerance),
                new XAttribute(nameof(PackageSize), PackageSize),
                new XAttribute(nameof(OpenType), OpenType),
                new XAttribute(nameof(ArrayType), ArrayType),
                new XAttribute(nameof(Temperature), Temperature),
                new XAttribute(nameof(Voltage), Voltage),
            };

            return CreateXmlElement(attributes);
        }
    }

    public partial class ChipCap
    {
        public static IEnumerable<ChipCap> Parse(ParsedExcelData excelData)
        {
            if (excelData.Context.Count == 0)
                yield break;

            foreach (var row in excelData.Context.Skip(1))
                yield return GetChipCap(row);
        }

        private static ChipCap GetChipCap(string[] row)
        {
            var capacitancePair = ConvertUnit(row[CapacitanceValueUnitIdx]);

            return new ChipCap
            {
                // 공통 property
                VendorCode = row[VendorCodeIdx],
                CustomerCode = row[CustomerCodeIdx],

                // 고유 속성 삽입
                Tolerance = ParseToDoubleOrDefault(row[ToleranceIdx], true),
                CapacitanceValue = capacitancePair.Value,
                CapacitanceUnit = ConvertSItoCapacitanceUnit(capacitancePair.Prefix),
                Voltage = ConvertUnit(row[VoltageIdx]).Value,
                // PackageSize 파싱 필요
                //PackageSize = ParsePackageSize(row[PackageSizeIdx]),
            };
        }

        //TODO: 수정중, PackageSize 파싱 보류 (작성 완료후 부모 클래스로 정의 옮기기)
        private static double ParsePackageSize(string packageSize)
        {
            //  ex "L=2.032 W=1.270" 받아서 2012로 반환
            if (packageSize == null || packageSize == string.Empty)
                throw new ArgumentException("", nameof(packageSize));

            string[] tokens = packageSize.Split(' '); // {"L=2.032", "W=1.270"}

            foreach (var curChar in tokens[0])
            {
                if (!Char.IsNumber(curChar)) continue;
            }

            return default;
        }
    }
}
