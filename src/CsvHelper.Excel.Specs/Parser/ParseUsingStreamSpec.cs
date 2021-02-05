using System.Globalization;
using System.IO;

namespace CsvHelper.Excel.Specs.Parser
{
    public class ParseUsingStreamSpec : ExcelParserSpec
    {
        public ParseUsingStreamSpec() : base("parse_by_stream")
        {
            using var stream = File.OpenRead(Path);
            using var parser = new ExcelParser(stream, CultureInfo.InvariantCulture);
            Run(parser);
        }
    }
}