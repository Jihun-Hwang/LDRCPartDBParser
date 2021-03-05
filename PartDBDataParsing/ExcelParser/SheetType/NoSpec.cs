using Pentacube.CAx.ECAD.LcadIf;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PartDBDataParsing.ExcelParser.SheetType
{
    public sealed class NoSpec : CommonExcelData
    {
        public NoSpec() : base(ePartClassification.Uncategorized)
        {
        }

        public override XElement CreateXmlElement()
        {
            return CreateXmlElement(new XAttribute[0]);
        }

        public static IEnumerable<NoSpec> Parse(ParsedExcelData excelData)
        {
            if (excelData.Context.Count == 0)
                yield break;

            foreach (var row in excelData.Context.Skip(1))
            {
                yield return new NoSpec
                {
                    VendorCode = row[VendorCodeIdx],
                    CustomerCode = row[CustomerCodeIdx],
                };
            }
        }
    }
}
