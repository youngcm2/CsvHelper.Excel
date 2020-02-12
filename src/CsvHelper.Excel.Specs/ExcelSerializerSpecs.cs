using System;
using System.IO;
using ClosedXML.Excel;
using FluentAssertions;
using Xunit;
using Xunit.Abstractions;

namespace CsvHelper.Excel.Specs
{
    public class ExcelSerializerSpecs
    {
        public abstract class Spec : IDisposable
        {
            protected readonly Person[] Values =
            {
                new Person {Name = "Bill", Age = 20},
                new Person {Name = "Ben", Age = 20},
                new Person {Name = "Weed", Age = 30}
            };

            protected string Path { get; }

            public ITestOutputHelper OutputHelper { get; private set; }
            protected string WorksheetName { get; }

            protected int StartRow { get; }

            protected int StartColumn { get; }

            protected abstract XLWorkbook GetWorkbook();
            protected abstract IXLWorksheet GetWorksheet();

            protected Spec(ITestOutputHelper outputHelper, string path, string worksheetName = "Export",
                int startRow = 1, int startColumn = 1)
            {
                Path =
                    System.IO.Path.GetFullPath(System.IO.Path.Combine("data", Guid.NewGuid().ToString(), $"{path}.xlsx"));

                var directory = System.IO.Path.GetDirectoryName(Path);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                OutputHelper = outputHelper;
                WorksheetName = worksheetName;
                StartRow = startRow;
                StartColumn = startColumn;
            }

            protected void Run(ExcelSerializer serialiser)
            {
                using var writer = new CsvWriter(serialiser);
                writer.Configuration.AutoMap<Person>();
                writer.WriteRecords(Values);
            }

            [Fact]
            public void TheFileIsAValidExcelFile()
            {
                GetWorkbook().Should().NotBeNull();
            }

            [Fact]
            public void TheExcelWorkbookHeadersAreCorrect()
            {
                nameof(Person.Name).Should().Be(GetWorksheet().Row(StartRow).Cell(StartColumn).Value.ToString());
                nameof(Person.Age).Should().Be(GetWorksheet().Row(StartRow).Cell(StartColumn + 1).Value.ToString());
            }

            [Fact]
            public void TheExcelWorkbookValuesAreCorrect()
            {
                for (var i = 0; i < Values.Length; i++)
                {
                    Values[i].Name.Should().Be(GetWorksheet().Row(StartRow + i + 1).Cell(StartColumn).Value.ToString());
                    Values[i].Age.ToString().Should().Be(
                        GetWorksheet().Row(StartRow + i + 1).Cell(StartColumn + 1).Value.ToString());
                }
            }

            public void Dispose()
            {
                GetWorkbook()?.Dispose();
                // Helpers.Delete(Path);
            }
        }

        public class SerialiseUsingPathSpec : Spec
        {
            public SerialiseUsingPathSpec(ITestOutputHelper outputHelper) : base(outputHelper, "serialise_by_path")
            {
                using var serialiser = new ExcelSerializer(Path);
                Run(serialiser);
            }

            protected override XLWorkbook GetWorkbook() => Helpers.GetOrCreateWorkbook(Path, WorksheetName);

            protected override IXLWorksheet GetWorksheet()
                => Helpers.GetOrCreateWorkbook(Path, WorksheetName).GetOrAddWorksheet(WorksheetName);
        }

        public class SerialiseUsingPathWithOffsetsSpec : Spec
        {
            public SerialiseUsingPathWithOffsetsSpec(ITestOutputHelper outputHelper)
                : base(outputHelper, "serialise_by_path_with_offsets", "Export", startRow: 5, startColumn: 5)
            {
                using var serialiser = new ExcelSerializer(Path)
                {
                    ColumnOffset = StartColumn - 1,
                    RowOffset = StartRow - 1
                };
                Run(serialiser);
            }

            protected override XLWorkbook GetWorkbook() => Helpers.GetOrCreateWorkbook(Path, WorksheetName);

            protected override IXLWorksheet GetWorksheet()
                => Helpers.GetOrCreateWorkbook(Path, WorksheetName).GetOrAddWorksheet(WorksheetName);
        }

        public class SerialiseUsingPathAndSheetnameSpec : Spec
        {
            public SerialiseUsingPathAndSheetnameSpec(ITestOutputHelper outputHelper)
                : base(outputHelper, $"serialise_by_path_and_sheetname", "a_different_sheet_name")
            {
                using var serialiser = new ExcelSerializer(Path, WorksheetName);
                Run(serialiser);
            }

            protected override XLWorkbook GetWorkbook() => Helpers.GetOrCreateWorkbook(Path, WorksheetName);

            protected override IXLWorksheet GetWorksheet()
                => Helpers.GetOrCreateWorkbook(Path, WorksheetName).GetOrAddWorksheet(WorksheetName);
        }

        public class SerialiseUsingStreamSpec : Spec
        {
            private readonly byte[] bytes;

            public SerialiseUsingStreamSpec(ITestOutputHelper outputHelper)
                : base(outputHelper, "serialise_by_workbook")
            {
                using var stream = new MemoryStream();
                using var serialiser = new ExcelSerializer(stream);
                Run(serialiser);

                bytes = stream.ToArray();
            }

            protected override XLWorkbook GetWorkbook()
            {
                using var stream = new MemoryStream(bytes);
                return new XLWorkbook(stream);
            }

            protected override IXLWorksheet GetWorksheet()
            {
                return GetWorkbook().GetOrAddWorksheet(WorksheetName);
            }
        }

        public class SerialiseUsingStreamAndSheetnameSpec : Spec
        {
            private readonly byte[] bytes;

            public SerialiseUsingStreamAndSheetnameSpec(ITestOutputHelper outputHelper)
                : base(outputHelper, "serialise_by_workbook_and_sheetname", "a_different_sheet_name")
            {
                using var stream = new MemoryStream();
                using var serialiser = new ExcelSerializer(stream, WorksheetName);
                Run(serialiser);

                bytes = stream.ToArray();
            }

            protected override XLWorkbook GetWorkbook()
            {
                using var stream = new MemoryStream(bytes);
                return new XLWorkbook(stream);
            }

            protected override IXLWorksheet GetWorksheet()
            {
                return GetWorkbook().GetOrAddWorksheet(WorksheetName);
            }
        }
    }
}