using ExcelDataReader;
using System.Collections.Generic;
using System.IO;
using Converter.ExcelConverter.Extensions;

namespace Converter.ExcelConverter.Services
{
    public interface IExcelConverterService
    {
        /// <summary>
        /// Convert an Excel file in Stream to a Dictionary.
        /// The inputted Excel file is supposed to be a data-only Excel, because we do not recognize graphs.
        /// </summary>
        /// <param name="fileStream">The Excel file in Stream</param>
        /// <param name="useExcelTypes">Use the Excel cell's type. Otherwise, all will be string</param>
        /// <param name="hasHeaders">Use the Excel first line as object name</param>
        /// <returns>The Excel file converted to a standard Dictionary</returns>
        Dictionary<string, ICollection<Dictionary<string, object>>> ConvertExcelToDictionary(Stream fileStream, bool useExcelTypes = true, bool hasHeaders = true);
    }

    public class ExcelConverterService : IExcelConverterService
    {
        public ExcelConverterService()
        {
            
        }

        /// <summary>
        /// Convert an Excel file in Stream to a Dictionary.
        /// The inputted Excel file is supposed to be a data-only Excel, because we do not recognize graphs.
        /// </summary>
        /// <param name="fileStream">The Excel file in Stream</param>
        /// <param name="useExcelTypes">Use the Excel cell's type. Otherwise, all will be string</param>
        /// <param name="hasHeaders">Use the Excel first line as object name</param>
        /// <returns>The Excel file converted to a standard Dictionary</returns>
        public Dictionary<string, ICollection<Dictionary<string, object>>> ConvertExcelToDictionary(Stream fileStream, bool useExcelTypes = true, bool hasHeaders = true)
        {
            using (var stream = fileStream)
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    return reader.AsDataSet().Tables.AsDictionary();
                }
            }
        }
    }
}
