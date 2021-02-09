using System.Globalization;
using System.IO;
using ClosedXML.Excel;
using CsvHelper.Excel.Specs.Common;
using Xunit.Abstractions;

namespace CsvHelper.Excel.Specs.Writer
{
    public class SerialiseUsingStreamAndSheetnameSpec : ExcelWriterSpec
    {
        private readonly byte[] _bytes;

        public SerialiseUsingStreamAndSheetnameSpec(ITestOutputHelper outputHelper)
            : base(outputHelper, "serialise_by_workbook_and_sheetname", "a_different_sheet_name")
        {
            using var stream = new MemoryStream();
            using (var excelWriter = new ExcelWriter(stream, WorksheetName, CultureInfo.InvariantCulture))
            {
                Run(excelWriter);
            }

            _bytes = stream.ToArray();
        }

        protected override XLWorkbook GetWorkbook()
        {
            using var stream = new MemoryStream(_bytes);
            return new XLWorkbook(stream);
        }

        protected override IXLWorksheet GetWorksheet()
        {
            return GetWorkbook().GetOrAddWorksheet(WorksheetName);
        }
    }
}