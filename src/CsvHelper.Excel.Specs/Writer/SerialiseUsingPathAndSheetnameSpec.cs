using System.Globalization;
using ClosedXML.Excel;
using CsvHelper.Excel.Specs.Common;
using Xunit.Abstractions;

namespace CsvHelper.Excel.Specs.Writer
{
    public class SerialiseUsingPathAndSheetnameSpec : ExcelWriterSpec
    {
        public SerialiseUsingPathAndSheetnameSpec(ITestOutputHelper outputHelper)
            : base(outputHelper, $"serialise_by_path_and_sheetname", "a_different_sheet_name")
        {
            using var excelWriter = new ExcelWriter(Path, WorksheetName, CultureInfo.InvariantCulture);
            Run(excelWriter);
        }

        protected override XLWorkbook GetWorkbook() => Helpers.GetOrCreateWorkbook(Path, WorksheetName);

        protected override IXLWorksheet GetWorksheet()
            => Helpers.GetOrCreateWorkbook(Path, WorksheetName).GetOrAddWorksheet(WorksheetName);
    }
}