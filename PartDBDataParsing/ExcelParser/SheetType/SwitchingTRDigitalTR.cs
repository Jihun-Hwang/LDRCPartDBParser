using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PartDBDataParsing.ExcelParser.SheetType
{
    /// <summary>
    /// Classification : 미정
    /// </summary>
    public sealed partial class SwitchingTRDigitalTR : CommonExcelData
    {
        // row index

        // 고유 속성

        // Pin

        public override XElement CreateXmlElement()
        {
            var attributes = new[]
            {
                new XAttribute("UniqueAttribute1", "Empty"),  // TODO: CLASSIFICATION미정
            };

            return CreateXmlElement(attributes);
        }
    }

    public partial class SwitchingTRDigitalTR
    {
        public static IEnumerable<SwitchingTRDigitalTR> Parse(ParsedExcelData excelData)
        {
            if (excelData.Context.Count == 0)
                yield break;

            foreach (var row in excelData.Context.Skip(1))
                yield return GetSwitchingTRDigitalTR(row);
        }

        private static SwitchingTRDigitalTR GetSwitchingTRDigitalTR(string[] row)
        {
            return new SwitchingTRDigitalTR();
        }
    }
}
