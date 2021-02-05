using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using CsvHelper.Excel.Specs.Common;
using FluentAssertions;
using Xunit;

namespace CsvHelper.Excel.Specs.Parser
{
    public abstract class EmptySpecWithNoHeaders : IDisposable
    {
        protected readonly Person[] Values = new Person[0];

        protected Person[] Results;

        protected string Path { get; }

        protected string WorksheetName { get; }

        protected int StartRow { get; }

        protected int StartColumn { get; }

        protected XLWorkbook Workbook { get; }

        protected IXLWorksheet Worksheet { get; }

        protected EmptySpecWithNoHeaders(string path, string worksheetName = "Export", int startRow = 1,
            int startColumn = 1)
        {
            Path =
                System.IO.Path.GetFullPath(
                    System.IO.Path.Combine("data", Guid.NewGuid().ToString(), $"{path}.xlsx"));

            var directory = System.IO.Path.GetDirectoryName(Path);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory!);
            }

            WorksheetName = worksheetName;
            StartRow = startRow;
            StartColumn = startColumn;
            Workbook = Helpers.GetOrCreateWorkbook(Path, WorksheetName);
            Worksheet = Workbook.GetOrAddWorksheet(WorksheetName);

            Workbook.SaveAs(Path);
        }

        protected void Run(ExcelParser parser)
        {
            using var reader = new CsvReader(parser);

            reader.Context.AutoMap<Person>();
            var records = reader.GetRecords<Person>();
            Results = records.ToArray();
        }

        [Fact]
        public void TheResultsAreNotNull()
        {
            Results.Should().NotBeNull();
        }

        [Fact]
        public void TheResultsAreCorrect()
        {
            Values.Should().BeEquivalentTo(Results, options => options.IncludingProperties());
        }

        public void Dispose()
        {
            Workbook?.Dispose();
            Helpers.Delete(Path);
        }
    }
}