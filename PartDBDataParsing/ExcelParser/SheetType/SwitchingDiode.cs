using Pentacube.CAx.ECAD.LcadIf;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PartDBDataParsing.ExcelParser.SheetType
{
    /// <summary>
    /// Classification : Diode
    /// </summary>
    public sealed partial class SwitchingDiode : CommonExcelData
    {
        // row index
        private const int AnodePinNumIdx = 10;
        private const int VoltageIdx = 15;

        // 고유 속성
        public string AnodePinNum { get; set; } = string.Empty;   // 애노드 핀번호 - o
        public double Voltage { get; set; }                       // 최대 역전압(V) - o
        public double Power { get; set; }                         // 정격 전력(W) - x
        public bool ArrayType { get; set; } = default;            // ArrayType - x

        public SwitchingDiode() : base(ePartClassification.Diode)
        {
        }
        public override XElement CreateXmlElement()
        {
            var attributes = new[]
            {
                new XAttribute(nameof(AnodePinNum), AnodePinNum),
                new XAttribute(nameof(Voltage), Voltage),
                new XAttribute(nameof(Power), Power),
                new XAttribute(nameof(ArrayType), ArrayType),
            };

            return CreateXmlElement(attributes);
        }
    }

    public partial class SwitchingDiode
    {
        public static IEnumerable<SwitchingDiode> Parse(ParsedExcelData excelData)
        {
            if (excelData.Context.Count == 0)
                yield break;

            foreach (var row in excelData.Context.Skip(1))
                yield return GetSwitchingDiode(row);
        }

        private static SwitchingDiode GetSwitchingDiode(string[] row)
        {
            return new SwitchingDiode
            {
                // 공통 property
                VendorCode = row[VendorCodeIdx],
                CustomerCode = row[CustomerCodeIdx],

                // 고유 속성 삽입
                //TODO: AnodePinNum 보류
                Voltage = ConvertUnit(row[VoltageIdx]).Value,

            };
        }
    }
}
