using Pentacube.Infrastructure.OfficeManager.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PartDBDataParsing.ExcelParser
{
    public class DefaultExcelParser
    {
        public static IEnumerable<ParsedExcelData> Parse(string path)
        {
            if (path == null)
                throw new ArgumentNullException(nameof(path), "path should not be Null");
            if (!File.Exists(path))   // path가 Empty인 경우도 포함
                throw new FileNotFoundException("File not found", nameof(path));

            using (var spread = SpreadBook.Open(path))
            {
                foreach (var sheet in spread.Sheets)
                {
                    var rows = GetTableRows(sheet);
                    if (rows.Count > 1)
                        yield return new ParsedExcelData(sheet.Name, rows);
                }
            }
        }

        private static List<string[]> GetTableRows(SpreadSheet sheet)
        {
            var headers = GetHeaders(sheet).ToArray();
            var rows = new List<string[]> { headers };

            int rowIdx = 2;
            while (true)
            {
                var cell = new SpreadCell(rowIdx++, 1);

                var row = new List<string>();
                for (int col = 0; col < headers.Length; col++)
                    row.Add(cell.MoveColumn(col)[sheet].Text);

                if (row.All(string.IsNullOrEmpty))
                    break;

                rows.Add(row.ToArray());
            }

            return rows;
        }

        private static List<string> GetHeaders(SpreadSheet sheet)
        {
            if (sheet == null)
                return Enumerable.Empty<string>().ToList();

            var headers = new List<string>();

            var cell = new SpreadCell(1, 1);

            while (true)
            {
                var text = cell[sheet].Text;
                if (text.Length == 0)
                    break;

                headers.Add(text);
                cell = cell.MoveColumn(1);
            }

            return headers;
        }
    }
}
