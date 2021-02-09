using System.Globalization;

namespace CsvHelper.Excel.Specs.Parser
{
    public class ParseUsingPathAndSheetNameAndCultureSpec : ExcelParserSpec
    {
        public ParseUsingPathAndSheetNameAndCultureSpec() : base("parse_by_path_and_sheetname_and_culture",
            "a_different_sheet_name")
        {
            using var parser = new ExcelParser(Path, WorksheetName, CultureInfo.InvariantCulture);
            Run(parser);
        }
    }
}