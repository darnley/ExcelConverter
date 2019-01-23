using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace Holbor.Converter.Excel.Extensions
{
    public static class DataTableCollectionExtensions
    {
        /// <summary>
        /// Process a sheet collection to transform it to a dictionary object
        /// </summary>
        /// <param name="dataTableCollection">The table collection to be transformed</param>
        /// <param name="useDataTableTypes">Do you want that the program respect Excel field type?</param>
        /// <returns>The converted dictionary</returns>
        public static Dictionary<string, ICollection<Dictionary<string, object>>> AsDictionary(this DataTableCollection dataTableCollection, bool useDataTableTypes = true, bool hasHeaders = true)
        {
            // First you have the sheet name (string)
            // Second you have the columns into sheet (Dictionary<string, ...>)
            // Third you have the values of each column (Dictionary<..., ICollection<Dictionary<string, string>>>)

            // It is using Concurrent type because of Thread Safety in Parallel execution
            ConcurrentDictionary<string, ICollection<Dictionary<string, object>>> dictionary = new ConcurrentDictionary<string, ICollection<Dictionary<string, object>>>();

            // It will work with all sheets in parallel
            Parallel.For(0, dataTableCollection.Count, (currentIndex) =>
            {
                // Get the reference of current sheet
                DataTable sheet = dataTableCollection[currentIndex];

                ICollection<Dictionary<string, object>> sheetData = ProcessDataTable(sheet, useDataTableTypes, hasHeaders);

                // Add the sheet values to the collection of sheets
                // It uses the TryAdd according to concurrent programming
                dictionary.TryAdd(sheet.TableName, sheetData);
            });

            // Convert the concurrent dictionary to a normal dictionary
            return dictionary.ToDictionary(p => p.Key, v => v.Value);
        }

        /// <summary>
        /// Transform a DataTable to a collection
        /// </summary>
        /// <param name="dataTable">The DataTable to be converted</param>
        /// <param name="useDataTableTypes">We will consider the DataTable types</param>
        /// <param name="hasHeaders">Respect the DataTable headers</param>
        /// <returns>The converted DataTable</returns>
        private static ICollection<Dictionary<string, object>> ProcessDataTable(DataTable dataTable, bool useDataTableTypes = true, bool hasHeaders = true)
        {
            ICollection<Dictionary<string, object>> temporary = new List<Dictionary<string, object>>();

            // Get the columns of the current sheet
            // It is important because we should know how many columns we have
            DataColumnCollection columns = dataTable.Columns;

            // Iterate each row in the current sheet
            foreach (DataRow row in dataTable.Rows)
            {
                // Instantiate a dictionary to store the data from each row as key<string> and value<string>
                var rowDictionary = new Dictionary<string, object>();

                // Get each column value per row
                for (int columnIndex = 0; columnIndex < columns.Count; columnIndex++)
                {
                    string columnName = dataTable.Rows[0][columnIndex].ToString();
                    object columnValue = Convert.ChangeType(row[columnIndex], useDataTableTypes ? columns[columnIndex].DataType : Type.GetType("string"));

                    rowDictionary.Add(columnName, columnValue);
                }

                // Add the row values to the collection of the current sheet
                temporary.Add(rowDictionary);
            }

            if (hasHeaders)
            {
                temporary = temporary.Skip(1).ToList();
            }

            return temporary;
        }
    }
}
