namespace CsvHelper.Excel.Specs.Parser
{
    public class ParseUsingPathSpecEmptyWithHeaders : EmptySpecWithHeaders
    {
        public ParseUsingPathSpecEmptyWithHeaders() : base("parse_by_path_empty_with_headers")
        {
            using var parser = new ExcelParser(Path);
            Run(parser);
        }
    }
}