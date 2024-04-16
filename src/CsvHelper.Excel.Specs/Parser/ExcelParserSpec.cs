using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using CsvHelper.Excel.Specs.Common;
using FluentAssertions;
using Xunit;

namespace CsvHelper.Excel.Specs.Parser
{
    public abstract class ExcelParserSpec : IDisposable
    {
        protected readonly Person[] Values =
        {
            new() {Name = "Bill", Age = 40},
            new() {Name = "Ben", Age = 30},
            new() {Name = "Weed", Age = 40}
        };

        protected Person[] Results;

        protected string Path { get; }

        protected string WorksheetName { get; }

        protected int StartRow { get; }

        protected int StartColumn { get; }

        protected XLWorkbook Workbook { get; }

        protected IXLWorksheet Worksheet { get; }

        protected ExcelParserSpec(string path, string worksheetName = "Export", int startRow = 1, int startColumn = 1,
            bool includeBlankRow = false)
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
            var currentRow = StartRow;

            var headerRow = Worksheet.Row(currentRow);
            headerRow.Cell(StartColumn).Value = nameof(Person.Name);
            headerRow.Cell(StartColumn + 1).Value = nameof(Person.Age);
            if (includeBlankRow)
            {
                currentRow++;
                Worksheet.Row(currentRow);
            }

            for (var i = 0; i < Values.Length; i++)
            {
                var row = Worksheet.Row(currentRow + i + 1);
                row.Cell(StartColumn).Value = Values[i].Name;
                row.Cell(StartColumn + 1).Value = Values[i].Age;
            }

            if (includeBlankRow)
            {
                currentRow++;
                Worksheet.Row(currentRow);
            }

            Workbook.SaveAs(Path);
        }

        protected void Run(ExcelParser parser)
        {
            using var reader = new CsvReader(parser);

            reader.Context.AutoMap<Person>();
            var records = reader.GetRecords<Person>().ToList();
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