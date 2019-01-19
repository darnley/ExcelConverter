﻿using ExcelDataReader;
using System.Collections.Generic;
using System.IO;
using Converter.ExcelConverter.Extensions;

namespace Converter.ExcelConverter.Services
{
    public interface IExcelConverterService
    {
        Dictionary<string, ICollection<Dictionary<string, object>>> ConvertExcelToDictionary(Stream fileStream, bool useExcelTypes = true, bool hasHeaders = true);
    }

    public class ExcelConverterService : IExcelConverterService
    {
        public ExcelConverterService()
        {
            
        }

        public Dictionary<string, ICollection<Dictionary<string, object>>> ConvertExcelToDictionary(Stream fileStream, bool useExcelTypes = true, bool hasHeaders = true)
        {
            using (var stream = fileStream)
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    return reader.AsDataSet().Tables.AsDictionary();
                }
            }
        }
    }
}