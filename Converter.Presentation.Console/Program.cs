using Converter.ExcelConverter.Services;
using System;
using System.IO;

namespace Converter.Presentation.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            ExcelConverterService service = new ExcelConverterService();

            using (var stream = File.Open("D:/Users/Darnley/Desktop/file_excel.xlsx", FileMode.Open, FileAccess.Read))
            {
                service.ConvertExcelToDictionary(stream);
            }
        }
    }
}
