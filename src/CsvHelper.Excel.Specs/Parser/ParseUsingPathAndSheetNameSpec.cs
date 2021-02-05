namespace CsvHelper.Excel.Specs.Parser
{
    public class ParseUsingPathAndSheetNameSpec : ExcelParserSpec
    {
        public ParseUsingPathAndSheetNameSpec() : base("parse_by_path_and_sheetname", "a_different_sheet_name")
        {
            using var parser = new ExcelParser(Path, WorksheetName);
            Run(parser);
        }
    }
}