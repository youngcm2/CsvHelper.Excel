using System.Globalization;
using System.Linq;
using CsvHelper.Configuration;

namespace CsvHelper.Excel.Specs.Parser
{
    public class ParseUsingPathSpecWithBlankRow : ExcelParserSpec
    {
        public ParseUsingPathSpecWithBlankRow() : base("parse_by_path_with_blank_row", includeBlankRow: true)
        {
            var csvConfiguration = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                ShouldSkipRecord = args => args.Row.Parser.Record?.All(string.IsNullOrWhiteSpace) ?? false,
            };
            using var parser = new ExcelParser(Path, null, csvConfiguration);
            Run(parser);
        }
    }
}