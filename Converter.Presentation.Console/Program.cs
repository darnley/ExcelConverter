using Converter.ExcelConverter.Services;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Diagnostics;

namespace Converter.Presentation.Console
{
    class Program
    {
        static int Main(string[] args)
        {
            System.Console.WriteLine($"[{DateTime.UtcNow.ToString("o")}] Processing...");

            Stopwatch operationTime = Stopwatch.StartNew();

            // Test if input arguments were supplied:
            if (args.Length == 0)
            {
                System.Console.WriteLine("Please enter the Microsoft Excel file path argument.");
                System.Console.WriteLine("Usage: <path>");
                return 1;
            }

            string excelFilePath = Path.GetFullPath(args[0]);

            ExcelConverterService service = new ExcelConverterService();

            using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
            {
                var result = service.ConvertExcelToDictionary(stream);

                string fileName = Path.GetFileNameWithoutExtension(excelFilePath);
                string fileDirectory = Path.GetDirectoryName(excelFilePath);

                string jsonFilePath = Path.Combine(fileDirectory, fileName + ".json");

                using (StreamWriter file = File.CreateText(jsonFilePath))
                {
                    JsonSerializer serializer = new JsonSerializer();
                    //serialize object directly into file stream
                    serializer.Serialize(file, result);
                }
            }

            operationTime.Stop();
            
            System.Console.WriteLine($"[{DateTime.UtcNow.ToString("o")}] Finished in {operationTime.ElapsedMilliseconds} milliseconds");

            return 0;
        }
    }
}
