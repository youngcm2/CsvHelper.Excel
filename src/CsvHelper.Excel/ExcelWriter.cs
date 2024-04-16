using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using ClosedXML.Excel;
using CsvHelper.Configuration;

#pragma warning disable 649
#pragma warning disable 169

namespace CsvHelper.Excel
{
    /// <summary>
    /// Used to write CSV files.
    /// </summary>
    public class ExcelWriter : CsvWriter
    {
        private readonly bool _leaveOpen;
        private readonly bool _sanitizeForInjection;

        private bool _disposed;
        private int _row = 1;
        private int _index = 1;
        private readonly IXLWorksheet _worksheet;
        private readonly Stream _stream;

        public override int Index => _index;
        public override int Row => _row;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        public ExcelWriter(string path) : this(File.Open(path, FileMode.OpenOrCreate, FileAccess.Write), "export",
            CultureInfo.InvariantCulture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="culture">The culture.</param>
        public ExcelWriter(string path, CultureInfo culture) : this(
            File.Open(path, FileMode.OpenOrCreate, FileAccess.Write), "export", culture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="sheetName">The sheet name</param>
        public ExcelWriter(string path, string sheetName) : this(
            File.Open(path, FileMode.OpenOrCreate, FileAccess.Write), sheetName, CultureInfo.InvariantCulture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="culture">The culture.</param>
        public ExcelWriter(string path, string sheetName, CultureInfo culture) : this(
            File.Open(path, FileMode.OpenOrCreate, FileAccess.Write), sheetName, culture)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="culture">The culture.</param>
        /// <param name="leaveOpen"><c>true</c> to leave the <see cref="TextWriter"/> open after the <see cref="ExcelWriter"/> object is disposed, otherwise <c>false</c>.</param>
        public ExcelWriter(Stream stream, CultureInfo culture, bool leaveOpen = false) : this(stream, "export", culture,
            leaveOpen)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="culture">The culture.</param>
        /// <param name="leaveOpen"><c>true</c> to leave the <see cref="TextWriter"/> open after the <see cref="ExcelWriter"/> object is disposed, otherwise <c>false</c>.</param>
        public ExcelWriter(Stream stream, string sheetName, CultureInfo culture, bool leaveOpen = false) : this(stream,
            sheetName, new CsvConfiguration(culture), leaveOpen)
        {
            
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWriter"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <param name="sheetName">The sheet name</param>
        /// <param name="configuration">The configuration.</param>
        private ExcelWriter(Stream stream, string sheetName, CsvConfiguration configuration, bool leaveOpen = false) : base(TextWriter.Null,
            configuration, leaveOpen)
        {
            configuration.Validate();
            _worksheet = new XLWorkbook().AddWorksheet(sheetName);
            this._stream = stream;

            _leaveOpen = leaveOpen;
            _sanitizeForInjection = configuration.InjectionOptions == InjectionOptions.Escape;
        }


        /// <inheritdoc/>
        public override void WriteField(string field, bool shouldQuote)
        {
            if (_sanitizeForInjection)
            {
                field = SanitizeForInjection(field);
            }

            WriteToCell(field);
            _index++;
        }

        /// <inheritdoc/>
        public override void NextRecord()
        {
            Flush();
            _index = 1;
            _row++;
        }

        /// <inheritdoc/>
        public override async Task NextRecordAsync()
        {
            await FlushAsync();
            _index = 1;
            _row++;
        }

        /// <inheritdoc/>
        public override void Flush()
        {
            _stream.Flush();
        }

        /// <inheritdoc/>
        public override Task FlushAsync()
        {
            return _stream.FlushAsync();
        }


        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void WriteToCell(string value)
        {
            var length = value?.Length ?? 0;

            if (value == null || length == 0)
            {
                return;
            }

            _worksheet.Worksheet.AsRange().Cell(_row, _index).Value = value;
        }

        /// <inheritdoc/>
        protected override void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            Flush();
            _worksheet.Workbook.SaveAs(_stream);
            _stream.Flush();

            if (disposing)
            {
                // Dispose managed state (managed objects)
                _worksheet.Workbook.Dispose();
                if (!_leaveOpen)
                {
                    _stream.Dispose();
                }
            }

            // Free unmanaged resources (unmanaged objects) and override finalizer
            // Set large fields to null

            _disposed = true;
        }

#if !NET45 && !NET47 && !NETSTANDARD2_0
		/// <inheritdoc/>
		protected override async ValueTask DisposeAsync(bool disposing)
		{
			if (_disposed)
			{
				return;
			}

			await FlushAsync().ConfigureAwait(false);
			_worksheet.Workbook.SaveAs(_stream);
			await _stream.FlushAsync().ConfigureAwait(false);

			if (disposing)
			{
				// Dispose managed state (managed objects)
				_worksheet.Workbook.Dispose();
				if (!_leaveOpen)
				{
					await _stream.DisposeAsync().ConfigureAwait(false);
				}
			}

			// Free unmanaged resources (unmanaged objects) and override finalizer
			// Set large fields to null


			_disposed = true;
		}
#endif
    }
}