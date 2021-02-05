using System.Globalization;
using ClosedXML.Excel;
using CsvHelper.Excel.Specs.Common;
using Xunit.Abstractions;

namespace CsvHelper.Excel.Specs.Writer
{
    public class SerialiseUsingPathSpec : ExcelWriterSpec
    {
        public SerialiseUsingPathSpec(ITestOutputHelper outputHelper) : base(outputHelper, "serialise_by_path")
        {
            using var excelWriter = new ExcelWriter(Path, CultureInfo.InvariantCulture);
            Run(excelWriter);
        }

        protected override XLWorkbook GetWorkbook() => Helpers.GetOrCreateWorkbook(Path, WorksheetName);

        protected override IXLWorksheet GetWorksheet()
            => Helpers.GetOrCreateWorkbook(Path, WorksheetName).GetOrAddWorksheet(WorksheetName);
    }
}