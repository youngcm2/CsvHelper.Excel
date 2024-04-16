using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using System;
using System.Linq;
using System.Runtime.CompilerServices;
using ClosedXML.Excel;
using CsvHelper.Configuration;

namespace CsvHelper.Excel
{
    /// <summary>
    /// Parses an Excel file.
    /// </summary>
    public class ExcelParser : IParser
    {
        private readonly bool _leaveOpen;

        private bool _disposed;
        private int _row = 1;
        private readonly IXLWorksheet _worksheet;
        private readonly Stream _stream;
        private int _rawRow = 1;
        private string[] _currentRecord;
        private readonly int _lastRow;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        public ExcelParser(string path) : this(
            File.Open(path, FileMode.OpenOrCreate, FileAccess.Read), null, CultureInfo.InvariantCulture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="sheetName">The sheet name</param>
        public ExcelParser(string path, string sheetName) : this(
            File.Open(path, FileMode.OpenOrCreate, FileAccess.Read), sheetName, CultureInfo.InvariantCulture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="culture">The culture.</param>
        public ExcelParser(string path, CultureInfo culture) : this(
            File.Open(path, FileMode.OpenOrCreate, FileAccess.Read), null, culture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="culture">The culture.</param>
        public ExcelParser(string path, string sheetName, CultureInfo culture) : this(
            File.Open(path, FileMode.OpenOrCreate, FileAccess.Read), sheetName, culture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="culture">The culture.</param>
        /// <param name="leaveOpen"><c>true</c> to leave the <see cref="TextWriter"/> open after the <see cref="ExcelParser"/> object is disposed, otherwise <c>false</c>.</param>
        public ExcelParser(Stream stream, CultureInfo culture, bool leaveOpen = false) : this(stream, null, culture,
            leaveOpen)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="culture">The culture.</param>
        /// <param name="leaveOpen"><c>true</c> to leave the <see cref="TextWriter"/> open after the <see cref="ExcelParser"/> object is disposed, otherwise <c>false</c>.</param>
        public ExcelParser(Stream stream, string sheetName, CultureInfo culture, bool leaveOpen = false) : this(stream,
            sheetName, new CsvConfiguration(culture), leaveOpen)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        /// <param name="path">The stream.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(string path, string sheetName, CsvConfiguration configuration) : this(
            File.Open(path, FileMode.OpenOrCreate, FileAccess.Read), sheetName, configuration)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelParser"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="configuration">The configuration.</param>
        public ExcelParser(Stream stream, string sheetName, CsvConfiguration configuration, bool leaveOpen = false)
        {
            var workbook = new XLWorkbook(stream);

            _worksheet = string.IsNullOrEmpty(sheetName) ? workbook.Worksheet(1) : workbook.Worksheet(sheetName);

            Configuration = configuration ?? new CsvConfiguration(CultureInfo.InvariantCulture);
            _stream = stream;
            var lastRowUsed = _worksheet.LastRowUsed();
            if (lastRowUsed != null)
            {
                _lastRow = lastRowUsed.RowNumber();

                var cellsUsed = _worksheet.CellsUsed();
                Count = cellsUsed.Max(c => c.Address.ColumnNumber) -
                    cellsUsed.Min(c => c.Address.ColumnNumber) + 1;
            }

            Context = new CsvContext(this);
            _leaveOpen = leaveOpen;
        }


        /// <inheritdoc/>
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing)
            {
                // Dispose managed state (managed objects)

                if (!_leaveOpen)
                {
                    _stream?.Dispose();
                }
            }

            // Free unmanaged resources (unmanaged objects) and override finalizer
            // Set large fields to null

            _disposed = true;
        }

        public bool Read()
        {
            if (Row > _lastRow)
            {
                return false;
            }

            _currentRecord = GetRecord();
            _row++;
            _rawRow++;
            return true;
        }

        public Task<bool> ReadAsync()
        {
            if (Row > _lastRow)
            {
                return Task.FromResult(false);
            }

            _currentRecord = GetRecord();
            _row++;
            _rawRow++;
            return Task.FromResult(true);
        }

        public long ByteCount => -1;
        public long CharCount => -1;
        public int Count { get; }

        public string this[int index] => Record.ElementAtOrDefault(index);

        public string[] Record => _currentRecord;

        public string RawRecord => string.Join(Configuration.Delimiter, Record);
        public int Row => _row;
        public int RawRow => _rawRow;

        public string Delimiter => Configuration.Delimiter;
        public CsvContext Context { get; }
        public IParserConfiguration Configuration { get; }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private string[] GetRecord()
        {
            var currentRow = _worksheet.Row(Row);
            var cells = currentRow.Cells(1, Count);
            var values = Configuration.TrimOptions.HasFlag(TrimOptions.Trim)
                ? cells.Select(x => x.Value.ToString()?.Trim()).ToArray()
                : cells.Select(x => x.Value.ToString()).ToArray();

            return values;
        }
    }
}