namespace CsvHelper.Excel.Specs.Parser
{
    public class ParseUsingPathSpec : ExcelParserSpec
    {
        public ParseUsingPathSpec() : base("parse_by_path")
        {
            using var parser = new ExcelParser(Path);
            Run(parser);
        }
    }
}