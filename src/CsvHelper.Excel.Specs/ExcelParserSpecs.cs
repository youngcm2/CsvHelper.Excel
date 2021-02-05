using System;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using CsvHelper.Configuration;
using FluentAssertions;
using Xunit;

namespace CsvHelper.Excel.Specs
{
    public class ExcelParserSpecs
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

            protected EmptySpecWithNoHeaders(string path, string worksheetName = "Export", int startRow = 1, int startColumn = 1)
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
        
        public class ParseUsingPathSpecEmptyWithNoHeaders : EmptySpecWithNoHeaders
        {
            public ParseUsingPathSpecEmptyWithNoHeaders() : base("parse_by_path_empty_with_no_headers")
            {
                using var parser = new ExcelParser(Path);
                Run(parser);
            }
        }
        
        public abstract class EmptySpecWithHeaders : IDisposable
        {
            protected readonly Person[] Values = new Person[0];

            protected Person[] Results;

            protected string Path { get; }

            protected string WorksheetName { get; }

            protected int StartRow { get; }

            protected int StartColumn { get; }

            protected XLWorkbook Workbook { get; }

            protected IXLWorksheet Worksheet { get; }

            protected EmptySpecWithHeaders(string path, string worksheetName = "Export", int startRow = 1, int startColumn = 1,
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
        
        public class ParseUsingPathSpecEmptyWithHeaders : EmptySpecWithHeaders
        {
            public ParseUsingPathSpecEmptyWithHeaders() : base("parse_by_path_empty_with_headers")
            {
                using var parser = new ExcelParser(Path);
                Run(parser);
            }
        }
        
        public abstract class Spec : IDisposable
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

            protected Spec(string path, string worksheetName = "Export", int startRow = 1, int startColumn = 1,
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
        
        public class ParseUsingPathSpec : Spec
        {
            public ParseUsingPathSpec() : base("parse_by_path")
            {
                using var parser = new ExcelParser(Path);
                Run(parser);
            }
        }

        public class ParseUsingPathSpecWithBlankRow : Spec
        {
            public ParseUsingPathSpecWithBlankRow() : base("parse_by_path_with_blank_row", includeBlankRow: true)
            {
                var csvConfiguration = new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    ShouldSkipRecord = record => record.All(string.IsNullOrEmpty)
                };
                using var parser = new ExcelParser(Path, null, csvConfiguration);
                Run(parser);
            }
        }

        public class ParseUsingPathSpecAndCulture : Spec
        {
            public ParseUsingPathSpecAndCulture() : base("parse_by_path_and_culture")
            {
                using var parser = new ExcelParser(Path, CultureInfo.InvariantCulture);
                Run(parser);
            }
        }

        public class ParseUsingPathAndSheetNameSpec : Spec
        {
            public ParseUsingPathAndSheetNameSpec() : base("parse_by_path_and_sheetname", "a_different_sheet_name")
            {
                using var parser = new ExcelParser(Path, WorksheetName);
                Run(parser);
            }
        }

        public class ParseUsingPathAndSheetNameAndCultureSpec : Spec
        {
            public ParseUsingPathAndSheetNameAndCultureSpec() : base("parse_by_path_and_sheetname_and_culture",
                "a_different_sheet_name")
            {
                using var parser = new ExcelParser(Path, WorksheetName, CultureInfo.InvariantCulture);
                Run(parser);
            }
        }

        public class ParseUsingStreamSpec : Spec
        {
            public ParseUsingStreamSpec() : base("parse_by_stream")
            {
                using var stream = File.OpenRead(Path);
                using var parser = new ExcelParser(stream, CultureInfo.InvariantCulture);
                Run(parser);
            }
        }

        public class ParseUsingStreamAndSheetNameSpec : Spec
        {
            public ParseUsingStreamAndSheetNameSpec() : base("parse_by_stream_and_sheetname", "a_different_sheet_name")
            {
                using var stream = File.OpenRead(Path);
                using var parser = new ExcelParser(stream, WorksheetName, CultureInfo.InvariantCulture);
                Run(parser);
            }
        }
    }
}