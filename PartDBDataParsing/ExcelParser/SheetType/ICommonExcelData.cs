using Pentacube.CAx.ECAD.LcadIf;

namespace PartDBDataParsing.ExcelParser.SheetType
{
    /// <summary>
    /// 모든 타입에서 공통으로 쓰이는 속성
    /// </summary>
    public interface ICommonExcelData
    {
        string VendorCode { get; }
        string CustomerCode { get; }
        ePartClassification Classification { get; }

        bool Aecq { get; }
        bool FailPart { get; }
        double MinOperationTemp { get; }
        double MaxOperationTemp { get; }
        string PinMap { get; }
        string SpiceModel { get; }
    }
}
