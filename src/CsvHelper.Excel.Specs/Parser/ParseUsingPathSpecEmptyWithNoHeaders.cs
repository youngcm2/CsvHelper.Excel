namespace CsvHelper.Excel.Specs.Parser
{
    public class ParseUsingPathSpecEmptyWithNoHeaders : EmptySpecWithNoHeaders
    {
        public ParseUsingPathSpecEmptyWithNoHeaders() : base("parse_by_path_empty_with_no_headers")
        {
            using var parser = new ExcelParser(Path);
            Run(parser);
        }
    }
}