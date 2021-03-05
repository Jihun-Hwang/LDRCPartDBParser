using System.Collections.Generic;

namespace PartDBDataParsing.ExcelParser
{
    public struct ParsedExcelData
    {
        public string SheetName { get; set; }
        public IReadOnlyList<string[]> Context { get; set; }

        public ParsedExcelData(string name, List<string[]> rows)
        {
            SheetName = name;
            Context = rows.AsReadOnly();
        }

        public override string ToString() => $"{SheetName} - Row count : {Context?.Count}";
    }
}
