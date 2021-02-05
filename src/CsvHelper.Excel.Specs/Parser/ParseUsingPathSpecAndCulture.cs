using System.Globalization;

namespace CsvHelper.Excel.Specs.Parser
{
    public class ParseUsingPathSpecAndCulture : ExcelParserSpec
    {
        public ParseUsingPathSpecAndCulture() : base("parse_by_path_and_culture")
        {
            using var parser = new ExcelParser(Path, CultureInfo.InvariantCulture);
            Run(parser);
        }
    }
}