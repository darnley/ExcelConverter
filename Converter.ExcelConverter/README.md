# Holbor.Converter.Excel
This package will help you to convert your [Microsoft Excel](https://products.office.com/excel) data-only files to a [Dictionary](https://docs.microsoft.com/dotnet/api/system.collections.generic.dictionary-2) and, if you would like that, convert to a [JSON](https://en.wikipedia.org/wiki/JSON), [XML](https://en.wikipedia.org/wiki/XML) or bind to your model.

## Usage
This package is ready for your dependency injector, because it already has an interface.
To use its service, use the `Holbor.Converter.Excel.Services.ExcelConverterService` class. Inside it, there is the `ConvertExcelToDictionary` method that you may use.

```csharp
using (var stream = File.Open("C:/file.xlsx", FileMode.Open, FileAccess.Read))
{
	var result = service.ConvertExcelToDictionary(stream);
}
```
If you would like to convert a `DataTableCollection` directly to a `Dictionary`, there is an [extension method](https://docs.microsoft.com/dotnet/csharp/programming-guide/classes-and-structs/extension-methods) which you can use. To use it, import the `Holbor.Converter.Excel.Extensions` and use the `.AsDictionary()` extension method.
```csharp
using (var dataset = randomObject.CreateDataSet(stream))
{
	return dataset.AsDictionary();
}
```
