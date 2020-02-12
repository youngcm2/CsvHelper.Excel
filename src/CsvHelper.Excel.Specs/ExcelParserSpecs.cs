using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using FluentAssertions;
using Xunit;

namespace CsvHelper.Excel.Specs
{
    public class ExcelParserSpecs
    {
        public abstract class Spec : IDisposable
        {
            protected readonly Person[] Values =
            {
                new Person {Name = "Bill", Age = 40},
                new Person {Name = "Ben", Age = 30},
                new Person {Name = "Weed", Age = 40}
            };

            protected Person[] Results;

            protected string Path { get; }

            protected string WorksheetName { get; }

            protected int StartRow { get; }

            protected int StartColumn { get; }

            protected XLWorkbook Workbook { get; }

            protected IXLWorksheet Worksheet { get; }

            protected Spec(string path, string worksheetName = "Export", int startRow = 1, int startColumn = 1)
            {
                Path =
                    System.IO.Path.GetFullPath(System.IO.Path.Combine("data", Guid.NewGuid().ToString(), $"{path}.xlsx"));
                
                var directory = System.IO.Path.GetDirectoryName(Path);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
                WorksheetName = worksheetName;
                StartRow = startRow;
                StartColumn = startColumn;
                Workbook = Helpers.GetOrCreateWorkbook(Path, WorksheetName);
                Worksheet = Workbook.GetOrAddWorksheet(WorksheetName);

                var headerRow = Worksheet.Row(StartRow);
                headerRow.Cell(StartColumn).Value = nameof(Person.Name);
                headerRow.Cell(StartColumn + 1).Value = nameof(Person.Age);
                for (var i = 0; i < Values.Length; i++)
                {
                    var row = Worksheet.Row(StartRow + i + 1);
                    row.Cell(StartColumn).Value = Values[i].Name;
                    row.Cell(StartColumn + 1).Value = Values[i].Age;
                }

                Workbook.SaveAs(Path);
            }

            protected void Run(ExcelParser parser)
            {
                using var reader = new CsvReader(parser);
                reader.Configuration.AutoMap<Person>();
                Results = reader.GetRecords<Person>().ToArray();
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

        public class ParseUsingPathWithOffsetsSpec : Spec
        {
            public ParseUsingPathWithOffsetsSpec() : base("parse_by_path_with_offset", "Export", 5, 5)
            {
                using var parser = new ExcelParser(Path) {ColumnOffset = StartColumn - 1, RowOffset = StartRow - 1};
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
        
        public class ParseUsingStreamSpec : Spec
        {
            public ParseUsingStreamSpec() : base("parse_by_stream")
            {
                using var stream = File.OpenRead(Path);
                using var parser = new ExcelParser(stream);
                Run(parser);
            }
        }

        public class ParseUsingStreamWithOffsetsSpec : Spec
        {
            public ParseUsingStreamWithOffsetsSpec() : base("parse_by_stream_with_offset", "Export", 5, 5)
            {
                using var stream = File.OpenRead(Path);
                using var parser = new ExcelParser(stream) {ColumnOffset = StartColumn - 1, RowOffset = StartRow - 1};
                Run(parser);
            }
        }

        public class ParseUsingStreamAndSheetNameSpec : Spec
        {
            public ParseUsingStreamAndSheetNameSpec() : base("parse_by_stream_and_sheetname", "a_different_sheet_name")
            {
                using var stream = File.OpenRead(Path);
                using var parser = new ExcelParser(stream, WorksheetName);
                Run(parser);
            }
        }
    }
}