using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using CsvHelper.Configuration;
using CsvHelper.Excel.Specs.Common;
using FluentAssertions;
using Xunit;

namespace CsvHelper.Excel.Specs.Parser
{
    public class ExcelParserAsyncSpec : IDisposable
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

        public ExcelParserAsyncSpec()
        {
            
            Path =
                System.IO.Path.GetFullPath(
                    System.IO.Path.Combine("data", Guid.NewGuid().ToString(), "excel_parser_async.xlsx"));

            var directory = System.IO.Path.GetDirectoryName(Path);
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory!);
            }

            WorksheetName = "worksheet_name";
            StartRow = 1;
            StartColumn = 1;
            Workbook = Helpers.GetOrCreateWorkbook(Path, WorksheetName);
            Worksheet = Workbook.GetOrAddWorksheet(WorksheetName);
            var currentRow = StartRow;

            var headerRow = Worksheet.Row(currentRow);
            headerRow.Cell(StartColumn).Value = nameof(Person.Name);
            headerRow.Cell(StartColumn + 1).Value = nameof(Person.Age);
            if (true)
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

            if (true)
            {
                currentRow++;
                Worksheet.Row(currentRow);
            }

            Workbook.SaveAs(Path);
        }

        protected async Task RunAsync()
        {
            var csvConfiguration = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                ShouldSkipRecord = x => x.Row.Parser.Record?.All(field => string.IsNullOrWhiteSpace(field)) ?? false
            };
            using var parser = new ExcelParser(Path, WorksheetName, csvConfiguration);
            using var reader = new CsvReader(parser);

            reader.Context.AutoMap<Person>();

            var records = reader.GetRecordsAsync<Person>();
            
            Results = await records.ToArrayAsync();
        }

        [Fact]
        public async void TheResultsAreNotNull()
        {
            await RunAsync();
            Results.Should().NotBeNull();
        }

        [Fact]
        public async void TheResultsAreCorrect()
        {
            await RunAsync();
            Values.Should().BeEquivalentTo(Results, options => options.IncludingProperties());
        }

        [Fact]
        public async void RowsHaveData()
        {
            var csvConfiguration = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                ShouldSkipRecord = x => x.Row.Parser.Record?.All(field => string.IsNullOrWhiteSpace(field)) ?? false
            };
            using var parser = new ExcelParser(Path, WorksheetName, csvConfiguration );
            using var reader = new CsvReader(parser);

            while (await reader.ReadAsync())
            {
                var data = reader.GetRecord<Person>();
                data.Should().NotBeNull();
            }
        }

        public void Dispose()
        {
            Workbook?.Dispose();
            Helpers.Delete(Path);
        }
    }
}