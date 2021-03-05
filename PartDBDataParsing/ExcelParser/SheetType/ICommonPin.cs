using PartDBDataParsing.ExcelParser.SheetType.PinType;
using System.Collections.Generic;

namespace PartDBDataParsing.ExcelParser.SheetType
{
    public interface ICommonPin
    {
        List<IPin> Pins { get; }
    }
}
