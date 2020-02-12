using ClosedXML.Excel;

namespace CsvHelper.Excel
{
    internal static class Helpers
    {
        public static IXLWorksheet GetOrAddWorksheet(this XLWorkbook workbook, string sheetName)
        {
            if (!workbook.TryGetWorksheet(sheetName, out var worksheet))
            {
                worksheet = workbook.AddWorksheet(sheetName);
            }
            return worksheet;
        }
    }
}