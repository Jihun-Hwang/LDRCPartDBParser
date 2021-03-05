using Pentacube.CAx.ECAD.LcadIf;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PartDBDataParsing.ExcelParser.SheetType
{
    /// <summary>
    /// Classification : SchottkyDiode
    /// </summary>
    public sealed partial class SchottkyDiode : CommonExcelData
    {
        // row index
        private const int AnodePinNumIdx = 10;
        private const int VoltageIdx = 18;

        // 고유 속성
        public string AnodePinNum { get; set; } = string.Empty;   // 애노드 핀번호 - o
        public double Voltage { get; set; }                       // 최대 역전압(A) - o
        public double Power { get; set; }                         // 정격 전력(W) - x

        public SchottkyDiode() : base(ePartClassification.SchottkyDiode)
        {
        }

        public override XElement CreateXmlElement()
        {
            var attributes = new[]
            {
                new XAttribute(nameof(AnodePinNum), AnodePinNum),
                new XAttribute(nameof(Voltage), Voltage),
                new XAttribute(nameof(Power), Power),
            };

            return CreateXmlElement(attributes);
        }
    }

    public partial class SchottkyDiode
    {
        public static IEnumerable<SchottkyDiode> Parse(ParsedExcelData excelData)
        {
            if (excelData.Context.Count == 0)
                yield break;

            foreach (var row in excelData.Context.Skip(1))
                yield return GetSchottkyDiode(row);
        }

        private static SchottkyDiode GetSchottkyDiode(string[] row)
        {
            return new SchottkyDiode
            {
                // 공통 property
                VendorCode = row[VendorCodeIdx],
                CustomerCode = row[CustomerCodeIdx],

                // 고유 속성 삽입
                // TODO: AnodePinNum 보류
                Voltage = ConvertUnit(row[VoltageIdx]).Value,
            };
        }
    }
}
