using System.Globalization;
using System.IO;
using ClosedXML.Excel;
using CsvHelper.Excel.Specs.Common;
using Xunit.Abstractions;

namespace CsvHelper.Excel.Specs.Writer
{
    public class SerialiseUsingStreamSpec : ExcelWriterSpec
    {
        private readonly byte[] _bytes;

        public SerialiseUsingStreamSpec(ITestOutputHelper outputHelper)
            : base(outputHelper, "serialise_by_workbook")
        {
            using var stream = new MemoryStream();
            using (var excelWriter = new ExcelWriter(stream, CultureInfo.InvariantCulture, true))
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