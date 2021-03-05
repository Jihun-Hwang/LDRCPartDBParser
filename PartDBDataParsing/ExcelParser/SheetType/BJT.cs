using PartDBDataParsing.ExcelParser.SheetType.PinType;
using Pentacube.CAx.ECAD.LcadIf;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace PartDBDataParsing.ExcelParser.SheetType
{
    /// <summary>
    /// Classification : BJT (ARRY TR, general purpose TR 시트)
    /// </summary>
    public sealed partial class BJT : CommonExcelData, ICommonPin
    {
        // row index
        private const int VcbIdx = 15;
        private const int VceIdx = 16;
        private const int VebIdx = 17;
        private const int CollectorCurrentIdx = 18;
        private const int PowerIdx = 20;
        private const int PinCellIdx = 22;

        // 고유 속성
        public eBJTType Type { get; set; } = default;             // 트랜지스터 유형 - x
        public double CollectorCurrent { get; set; }              // Collector 전류(A) - o
        public double Vce { get; set; }                           // 최대 Vce 전압- o
        public double Vcb { get; set; }                           // 최대 Vcb 전압 - o
        public double Veb { get; set; }                           // 최대 Veb 전압- o
        public double Resistance1 { get; set; }                   // 저항 1 - x
        public double Resistance2 { get; set; }                   // 저항 2 - x
        public bool ArrayType { get; set; } = default;            // ArrayType - x
        public double Power { get; set; }                         // 정격 전력 - o

        public List<IPin> Pins { get; } = new List<IPin>();

        public BJT() : base(ePartClassification.Bjt)
        {
        }

        public override XElement CreateXmlElement()
        {
            var attributes = new[]
            {
                new XAttribute(nameof(Type), Type),
                new XAttribute(nameof(CollectorCurrent), CollectorCurrent),
                new XAttribute(nameof(Vce), Vce),
                new XAttribute(nameof(Vcb), Vcb),
                new XAttribute(nameof(Veb), Veb),
                new XAttribute(nameof(Resistance1), Resistance1),
                new XAttribute(nameof(Resistance2), Resistance2),
                new XAttribute(nameof(ArrayType), ArrayType),
                new XAttribute(nameof(Power), Power),
            };

            return CreateXmlElement(attributes);
        }
    }

    public partial class BJT
    {
        public static IEnumerable<BJT> Parse(ParsedExcelData excelData)
        {
            if (excelData.Context.Count == 0)
                yield break;

            foreach (var row in excelData.Context.Skip(1))
                yield return GetBJT(row);
        }

        private static BJT GetBJT(string[] row)
        {
            var bjt = new BJT
            {
                // 공통 property
                VendorCode = row[VendorCodeIdx],
                CustomerCode = row[CustomerCodeIdx],

                // 고유 속성 삽입
                Vcb = ConvertUnit(row[VcbIdx]).Value,
                Vce = ConvertUnit(row[VceIdx]).Value,
                Veb = ConvertUnit(row[VebIdx]).Value,
                CollectorCurrent = ConvertUnit(row[CollectorCurrentIdx]).Value,
                Power = ConvertUnit(row[PowerIdx]).Value,
            };

            bjt.Pins.AddRange(GetPins(row[PinCellIdx]));

            return bjt;
        }

        private static IEnumerable<BJTPin> GetPins(string pinCell)
        {
            if (string.IsNullOrEmpty(pinCell))
                yield break;

            var convertedItems = pinCell.Split('-')
                .Select((str, idx) => (Num: $"{idx + 1}", Type: GetBJTPinType(str)))
                .Where(pair => pair.Type.HasValue)
                .Select(pair => (pair.Num, pair.Type.GetValueOrDefault()));

            foreach (var item in convertedItems)
                yield return new BJTPin(item);
/*
            if (string.IsNullOrEmpty(pinCell))
                return Enumerable.Empty<BJTPin>();

            return pinCell.Split('-')
                .Select((str, idx) => (Num: $"{idx + 1}", Type: GetBJTPinType(str)))
                .Where(pair => pair.Type.HasValue)
                .Select(pair => new BJTPin(pair.Num, pair.Type.GetValueOrDefault()));
*/
        }

        private static eBJTPinType? GetBJTPinType(string bjtPinTypeStr)
        {
            switch (bjtPinTypeStr.Trim().ToUpper())
            {
                case nameof(eBJTPinType.BASE): return eBJTPinType.BASE;
                case nameof(eBJTPinType.COLLECTOR): return eBJTPinType.COLLECTOR;
                case nameof(eBJTPinType.EMITTER): return eBJTPinType.EMITTER;

                default: return null;
            }
        }
    }
}
