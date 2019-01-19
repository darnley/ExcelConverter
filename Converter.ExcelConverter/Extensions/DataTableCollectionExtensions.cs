using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace Converter.ExcelConverter.Extensions
{
    public static class DataTableCollectionExtensions
    {
        private static Dictionary<string, ICollection<Dictionary<string, object>>> Initialize(this Dictionary<string, ICollection<Dictionary<string, object>>> dataTableCollection, long quantity)
        {
            for (long i = 0; i < quantity; i++)
            {
                dataTableCollection.Add(null, null);
            }

            return dataTableCollection;
        }

        /// <summary>
        /// Process a sheet collection to transform it to a dictionary object
        /// </summary>
        /// <param name="dataTableCollection">The table collection to be transformed</param>
        /// <param name="useDataTableTypes">Do you want that the program respect Excel field type?</param>
        /// <returns>The converted dictionary</returns>
        public static Dictionary<string, ICollection<Dictionary<string, object>>> AsDictionary(this DataTableCollection dataTableCollection, bool useDataTableTypes = true)
        {
            // First you have the sheet name (string)
            // Second you have the columns into sheet (Dictionary<string, ...>)
            // Third you have the values of each column (Dictionary<..., ICollection<Dictionary<string, string>>>)

            // It is using Concurrent type because of Thread Safety in Parallel execution
            ConcurrentDictionary<string, ICollection<Dictionary<string, object>>> dictionary = new ConcurrentDictionary<string, ICollection<Dictionary<string, object>>>();

            // It will work with all sheets in parallel
            Parallel.For(0, dataTableCollection.Count, (currentIndex) =>
            {
                ICollection<Dictionary<string, object>> temporary = new List<Dictionary<string, object>>();

                // Get the reference of current sheet
                DataTable sheet = dataTableCollection[currentIndex];

                // Get the columns of the current sheet
                // It is important because we should know how many columns we have
                DataColumnCollection columns = sheet.Columns;

                // Iterate each row in the current sheet
                foreach (DataRow row in sheet.Rows)
                {
                    // Instantiate a dictionary to store the data from each row as key<string> and value<string>
                    var rowDictionary = new Dictionary<string, object>();

                    // Get each column value per row
                    for (int columnIndex = 0; columnIndex < columns.Count; columnIndex++)
                    {
                        string columnName = sheet.Rows[0][columnIndex].ToString();
                        object columnValue = Convert.ChangeType(row[columnIndex], useDataTableTypes ? columns[columnIndex].DataType : Type.GetType("string"));

                        rowDictionary.Add(columnName, columnValue);
                    }

                    // Add the row values to the collection of the current sheet
                    temporary.Add(rowDictionary);
                }

                // Add the sheet values to the collection of sheets
                // It uses the TryAdd according to concurrent programming
                dictionary.TryAdd(sheet.TableName, temporary.Skip(1).ToList());
            });

            // Convert the concurrent dictionary to a normal dictionary
            return dictionary.ToDictionary(p => p.Key, v => v.Value);
        }
    }
}
