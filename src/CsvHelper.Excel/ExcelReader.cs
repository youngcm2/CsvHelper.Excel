using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using CsvHelper.Configuration;

namespace CsvHelper.Excel
{
    /// <summary>
    /// Read an Excel file.
    /// </summary>
    public class ExcelReader : CsvReader
    {

        public new ExcelParser Parser => (ExcelParser) base.Parser;

        /// <summary>Initializes a new instance of the <see cref="ExcelReader" /> class.</summary>
        /// <param name="path">The path.</param>
        public ExcelReader(string path) : base(new ExcelParser(path))
        {
        }

        /// <summary>Initializes a new instance of the <see cref="ExcelReader" /> class.</summary>
        /// <param name="path">The path.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        public ExcelReader(string path, string sheetName) : base(new ExcelParser(path, sheetName))
        {
        }

        /// <summary>Initializes a new instance of the <see cref="ExcelReader" /> class.</summary>
        /// <param name="path">The path.</param>
        /// <param name="culture">The culture.</param>
        public ExcelReader(string path, CultureInfo culture) : base(new ExcelParser(path, culture))
        {
        }

        /// <summary>Initializes a new instance of the <see cref="ExcelReader" /> class.</summary>
        /// <param name="path">The path.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="culture">The culture.</param>
        public ExcelReader(string path, string sheetName, CultureInfo culture) : base(new ExcelParser(path, sheetName, culture))
        {
        }

        /// <summary>Initializes a new instance of the <see cref="ExcelReader" /> class.</summary>
        /// <param name="stream">The stream.</param>
        /// <param name="culture">The culture.</param>
        public ExcelReader(Stream stream, CultureInfo culture) : base(new ExcelParser(stream, culture))
        {
        }

        /// <summary>Initializes a new instance of the <see cref="ExcelReader" /> class.</summary>
        /// <param name="stream">The stream.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="culture">The culture.</param>
        public ExcelReader(Stream stream, string sheetName, CultureInfo culture) : base(new ExcelParser(stream, sheetName, culture))
        {
        }

        /// <summary>Initializes a new instance of the <see cref="ExcelReader" /> class.</summary>
        /// <param name="parser">The Excel parser.</param>
        public ExcelReader(ExcelParser parser) : base(parser)
        {
        }

        /// <summary>Gets the comment on the specified column index, on current parsing row.</summary>
        /// <param name="index">The column index.</param>
        /// <returns>
        ///   <br />
        /// </returns>
        public string GetExcelComment(int index)
        {
            return Parser.GetComment(index);
        }

        /// <summary>Gets the comment on specific cell in the excel file.</summary>
        /// <param name="column">The column.</param>
        /// <param name="row">The row.</param>
        /// <returns>
        ///   <br />
        /// </returns>
        public string GetCommentAt(int column, int row)
        {
            return Parser.GetComment(column, row);
        }
    }
}
