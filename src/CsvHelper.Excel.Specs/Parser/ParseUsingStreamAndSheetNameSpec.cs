using System.Globalization;
using System.IO;

namespace CsvHelper.Excel.Specs.Parser
{
    public class ParseUsingStreamAndSheetNameSpec : ExcelParserSpec
    {
        public ParseUsingStreamAndSheetNameSpec() : base("parse_by_stream_and_sheetname", "a_different_sheet_name")
        {
            using var stream = File.OpenRead(Path);
            using var parser = new ExcelParser(stream, WorksheetName, CultureInfo.InvariantCulture);
            Run(parser);
        }
    }
}