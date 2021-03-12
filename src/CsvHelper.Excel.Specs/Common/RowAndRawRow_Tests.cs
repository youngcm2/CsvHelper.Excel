using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper.Configuration;
using FluentAssertions;
using Xunit;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace CsvHelper.Excel.Specs.Common
{
    public class DebugOutputHelper : ITestOutputHelper
    {
        public void WriteLine(string message)
        {
            Debug.WriteLine(message);
        }

        public void WriteLine(string format, params object[] args)
        {
            Debug.WriteLine(string.Format(format, args));
        }
    }
    public sealed class RowMap : ClassMap<MappingRow>
    {
        public RowMap()
        {
            this.Map(m => m.ResourceId)
                .Name("Resource Id");

            this.Map(m => m.ProductId)
                .Name("Product Id");
        }
    }

    public class MappingRow
    {
        public long? ResourceId { get; set; }

        public int? ProductId { get; set; }
    }
    public class RowAndRawRow_Tests
    {
        private readonly ITestOutputHelper _testOutputHelper;

        public RowAndRawRow_Tests()
        {
            _testOutputHelper = new DebugOutputHelper();

        }

        [Fact]
        public void RowReadWithExcelParser()
        {
            var bytes = File.ReadAllBytes(@"demo.xlsx");

            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                ShouldSkipRecord = record => record.Record.All(string.IsNullOrEmpty),
                HeaderValidated = args =>
                {
                    _testOutputHelper.WriteLine($"HeaderValidated: RawRow: {args.Context.Parser.RawRow}");
                }
            };
            var maxRow = 0;
            using (var stream = new MemoryStream(bytes))
            using (var parser = new ExcelParser(stream, "Asset Mapping Import", config))
            using (var reader = new CsvReader(parser))
            {
                reader.Context.RegisterClassMap<RowMap>();

                while (reader.Read())
                {
                    var dataRow = reader.GetRecord<MappingRow>();
                    if (reader.Context.Parser.RawRow > maxRow) maxRow = reader.Context.Parser.RawRow;
                    _testOutputHelper.WriteLine($"Data: ({dataRow.ProductId}, {dataRow.ResourceId}): RawRow: {reader.Context.Parser.RawRow}");
                }
            }

            maxRow.Should().Be(3);
        }

        [Fact]
        public void RowReadWithCsvParser()
        {
            var bytes = File.ReadAllBytes(@"demo.csv");

            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                ShouldSkipRecord = record => record.Record.All(string.IsNullOrEmpty),
                HeaderValidated = args =>
                {
                    _testOutputHelper.WriteLine($"HeaderValidated: RawRow: {args.Context.Parser.RawRow}");
                }
            };
            var maxRow = 0;
            using (var stream = new MemoryStream(bytes))
            using (var parser = new CsvParser(new StreamReader(stream,Encoding.UTF8),config))
            using (var reader = new CsvReader(parser))
            {
                reader.Context.RegisterClassMap<RowMap>();

                while (reader.Read())
                {
                    var dataRow = reader.GetRecord<MappingRow>();
                    if (reader.Context.Parser.RawRow > maxRow) maxRow = reader.Context.Parser.RawRow;
                    _testOutputHelper.WriteLine($"Data: ({dataRow.ProductId}, {dataRow.ResourceId}): RawRow: {reader.Context.Parser.RawRow}");
                }
            }
            maxRow.Should().Be(3);
        }
    }
}
