<!-- [![Build status](https://ci.appveyor.com/api/projects/status/bqh412kdla4peqsw?svg=true)](https://ci.appveyor.com/project/christophano/csvhelper-excel) -->

[![Discord Chat](https://img.shields.io/discord/308323056592486420.svg)](https://discord.gg/TbanjBb)  [![NuGet Badge](https://buildstats.info/nuget/CsvHelper.Excel.Core)](https://www.nuget.org/packages/CsvHelper.Excel.Core/)

# Csv Helper for Excel

CsvHelper for Excel is an extension that links 2 excellent libraries, [CsvHelper](https://github.com/JoshClose/CsvHelper) and [ClosedXml](https://github.com/closedxml/closedxml).
It provides an implementation of `ICsvParser` and `ICsvSerializer` from [CsvHelper](https://github.com/JoshClose/CsvHelper) that read and write to Excel using [ClosedXml](https://github.com/closedxml/closedxml).

### ExcelParser
`ExcelParser` implements `IParser` and allows you to specify the path of the workbook or a stream.

When the path is passed to the constructor then the workbook loading and disposal is handled by the parser. By default the first worksheet is used as the data source.
```csharp
using var reader = new CsvReader(new ExcelParser("path/to/file.xlsx"));
var people = reader.GetRecords<Person>();

```
When an instance of `stream` is passed to the constructor then disposal will not be handled by the parser. By default the first worksheet is used as the data source.
```csharp

var bytes = File.ReadAllBytes("path/to/file.xlsx");

using var stream = new MemoryStream(bytes);
using var parser = new ExcelParser(stream);
using var reader = new CsvReader(parser);

var people = reader.GetRecords<Person>();
// do other stuff with workbook

```

All constructor options have overloads allowing you to specify your own `CsvConfiguration`, otherwise the default is used.

### ExcelSerializer
`ExcelSerializer` implements `ISerializer` and, like `ExcelParser`, allows you to specify the path to which to (eventually) save the workbook or a stream.

When the path is passed to the constructor the creation and disposal of both the workbook and worksheet (defaultly named "Export") as well as the saving of the workbook on dispose, is handled by the serialiser.
```csharp
using (var writer = new CsvWriter(new ExcelSerializer("path/to/file.xlsx")))
{
    writer.WriteRecords(people);
}
```
When an instance of `stream` is passed to the constructor the creation and disposal of a new worksheet (defaultly named "Export") is handled by the serialiser, but the workbook will not be saved.
```csharp

using var stream = new MemoryStream();
using var serialiser = new ExcelParser(stream);
using var writer = new CsvWriter(serialiser);
writer.WriteRecords(people);
//other stuff
var bytes = stream.ToArray();
```
All constructor options have overloads allowing you to specify your own `CsvConfiguration`, otherwise the default is used.
