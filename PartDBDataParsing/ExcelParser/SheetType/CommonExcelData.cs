using JetBrains.Annotations;
using PartDBDataParsing.ExcelParser.SheetType.PinType;
using Pentacube.CAx.ECAD.LcadIf;
using Pentacube.Infrastructure.Service;
using System;
using System.Linq;
using System.Xml.Linq;

namespace PartDBDataParsing.ExcelParser.SheetType
{
    public abstract class CommonExcelData : ICommonExcelData
    {
        // excel column 공통 index
        public const int VendorCodeIdx = 0;
        public const int CustomerCodeIdx = 2;

        // 공통 attribute
        public string VendorCode { get; protected set; } = string.Empty;
        public string CustomerCode { get; protected set; } = string.Empty;
        public ePartClassification Classification { get; protected set; }

        public bool Aecq { get; protected set; } = default;
        public bool FailPart { get; protected set; } = default;
        public double MinOperationTemp { get; protected set; } = default;
        public double MaxOperationTemp { get; protected set; } = default;
        public string PinMap { get; protected set; } = string.Empty;
        public string SpiceModel { get; protected set; } = string.Empty;

        #region Constructor
        public CommonExcelData() : this(ePartClassification.None)
        {
        }

        public CommonExcelData(ePartClassification classification)
        {
            Classification = classification;
        }
        #endregion

        [NotNull]
        private XElement CreateXmlElementBase([NotNull, ItemNotNull] XAttribute[] attributes)
        {
            var part = new XElement(
                "Part",
                new XAttribute(nameof(VendorCode), VendorCode),
                new XAttribute(nameof(CustomerCode), CustomerCode),
                new XAttribute(nameof(Classification), Classification),
                new XAttribute(nameof(Aecq), Aecq),
                new XAttribute(nameof(FailPart), FailPart),
                new XAttribute(nameof(MinOperationTemp), MinOperationTemp),
                new XAttribute(nameof(MaxOperationTemp), MaxOperationTemp),
                new XAttribute(nameof(PinMap), PinMap),
                new XAttribute(nameof(SpiceModel), SpiceModel));

            if (attributes.Length > 0)
            {
                foreach (var attr in attributes)
                    part.Add(attr);
            }

            return part;
        }

        [NotNull]
        protected XElement CreateXmlElement([NotNull, ItemNotNull] XAttribute[] attributes)
        {
            var part = CreateXmlElementBase(attributes);

            if (this is ICommonPin pinPart)
            {
                foreach (var pin in pinPart.Pins.OfType<IPinSerialize>())
                    part.Add(pin.CreatePinXmlElement());
            }

            return part;
        }

        [NotNull]
        public abstract XElement CreateXmlElement();

        #region Static Functions
        protected static bool ParseToBoolOrDefault(string value)
        {
            return bool.TryParse(value, out var res) ? res : default;
        }

        protected static double ParseToDoubleOrDefault(string value, bool applyAbs = false)
        {
            double.TryParse(value, out var res);
            return applyAbs ? Math.Abs(res) : res;
        }

        public static (double Value, eSiPrefix Prefix) ConvertUnit(string value)
        {
            SiPrefixes.TryParseDouble(value, out var result, out var prefix);
            return (result, prefix);
        }

        /// <summary>
        /// eSiPrefix에서 eCapacitanceUnit로 반환
        /// </summary>
        protected static eCapacitanceUnit ConvertSItoCapacitanceUnit(eSiPrefix from)
        {
            switch (from)
            {
                case eSiPrefix.m: return eCapacitanceUnit.mF;
                case eSiPrefix.u: return eCapacitanceUnit.uF;
                case eSiPrefix.n: return eCapacitanceUnit.nF;
                case eSiPrefix.p: return eCapacitanceUnit.pF;

                default: return eCapacitanceUnit.F;
            }
        }

        /// <summary>
        /// eSiPrefix에서 eResistanceUnit로 반환
        /// </summary>
        protected static eResistanceUnit ConvertSItoResistanceUnit(eSiPrefix from)
        {
            switch (from)
            {
                case eSiPrefix.k: return eResistanceUnit.KOHM;
                case eSiPrefix.M: return eResistanceUnit.MOHM;

                default: return eResistanceUnit.OHM;
            }
        }

        /// <summary>
        /// eSiPrefix에서 eInductorUnit로 반환
        /// </summary>
        protected static eInductorUnit ConvertSItoInductorUnit(eSiPrefix from)
        {
            switch (from)
            {
                case eSiPrefix.m: return eInductorUnit.mH;
                case eSiPrefix.u: return eInductorUnit.uH;
                case eSiPrefix.n: return eInductorUnit.nH;
                case eSiPrefix.p: return eInductorUnit.pH;

                default: return eInductorUnit.H;
            }
        }
        #endregion Static Functions
    }
}
