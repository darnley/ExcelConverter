# ExcelConverter

This package allows you to convert [Microsoft Excel](https://products.office.com/excel) files to a Dictionary type. So, with that, you may have a JSON code.

## Installation

The main package that have the code is [.NET Standard 2.0](https://docs.microsoft.com/dotnet/standard/net-standard). So, it is compatible with many versions of [.NET Framework and .NET Core](https://dotnet.microsoft.com/download).

## Usage

Use the method `Converter.ExcelConverter.Services.ConvertExcelToDictionary()`. It receives a `Stream` that must contains the Microsoft Excel file (XLSX).

```csharp
using (var stream = System.IO.File.Open("C:/file.xlsx", FileMode.Open, FileAccess.Read))
{
    var dictionary = service.ConvertExcelToDictionary(stream);
    // Dictionary<string, ICollection<Dictionary<string, object>>>
}
```

If you would like to use it directly into a `DataTableCollection` object, use the [extension method](https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/classes-and-structs/extension-methods) `Converter.ExcelConverter.Extensions.AsDictionary()`. In this case, once you are using the [extension method](https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/classes-and-structs/extension-methods), you just have to call it as a `DataTableCollection` method.

```csharp
var result = dataSet.Tables.AsDictionary();
// Dictionary<string, ICollection<Dictionary<string, object>>>
```

#### Calling from Web API

By default, there is a basic Web API project you can use to send the Microsoft Excel file and receive the JSON code in response's body.

To do this, make a call as the following.

```
GET /api/converter HTTP/1.1
Host: localhost:5001
Content-Type: multipart/form-data;

Content-Disposition: form-data; name="file"; filename="C:\file_excel.xlsx"
```

Pay attention to the `Content-Disposition` name. The default request method uses the `file` parameter; so, you should put your [Microsoft Excel](https://products.office.com/excel) file into it.

## Compatibility

In the version 1.0, this package only supports XLSX files.

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License
[MIT](https://darnley.mit-license.org/2019)
