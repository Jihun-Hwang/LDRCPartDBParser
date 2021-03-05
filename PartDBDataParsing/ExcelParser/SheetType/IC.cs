using PartDBDataParsing.ExcelParser.SheetType.PinType;
using Pentacube.CAx.ECAD.LcadIf;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PartDBDataParsing.ExcelParser.SheetType
{
    /// <summary>
    /// Classification : IC (OPAMP, Regulator, DCDCPMIC 시트)
    /// </summary>
    public sealed partial class IC : CommonExcelData, ICommonPin
    {
        // row index

        // 고유 속성
        public eICType ICType { get; set; } = default;                    // IC 유형 - o
        public double CapacitanceValue { get; set; }                      // CAN IC 정전 용량 값 - x
        public eCapacitanceUnit CapacitanceUnit { get; set; } = default;  // CAN IC 정전 용량 단위 - x

        // PIN
        public List<IPin> Pins { get; } = new List<IPin>();

        public IC() : base(ePartClassification.Ic)
        {
        }

        public override XElement CreateXmlElement()
        {
            var attributes = new[]
            {
                new XAttribute(nameof(ICType), ICType),
                new XAttribute(nameof(CapacitanceValue), CapacitanceValue),
                new XAttribute(nameof(CapacitanceUnit), CapacitanceUnit),
            };

            return CreateXmlElement(attributes);
        }
    }

    public partial class IC
    {
        public static IEnumerable<IC> Parse(ParsedExcelData excelData)
        {
            if (excelData.Context.Count == 0)
                yield break;

            foreach (var row in excelData.Context.Skip(1))
                yield return GetIC(row, excelData.SheetName);
        }

        private static IC GetIC(string[] row, string sheetName)
        {
            return new IC
            {
                // 공통 property
                VendorCode = row[VendorCodeIdx],
                CustomerCode = row[CustomerCodeIdx],

                // 고유 속성 삽입
                ICType = GetICType(sheetName), // sheet name에 따라 ICType결정
            };
        }

        private static eICType GetICType(string sheetName)
        {
            switch (sheetName.ToUpper())
            {
                case "OP AMP": return eICType.OPAMP;
                case "REGULATOR": return eICType.REGULATOR;
                default: return default;   // TODO: eICType의 default값 알아내기
            }
        }
    }
}
