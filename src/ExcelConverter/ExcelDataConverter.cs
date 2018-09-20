using ExcelConverter.Crosscutting;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelConverter
{
    public class ExcelDataConverter
    {
        public static IEnumerable<T> Get<T>(Stream stream, bool hasHeader = true)
        {
            var result = new List<T>();
            var props = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public).Where(x => x.CanWrite)
                .ToArray();

            var table = stream.ToDataTable();

            if (table.Columns.Count != props.Length)
                throw new Exception("Number of columns doesn't match number of properties");

            if (hasHeader) table.Rows.RemoveAt(0);

            foreach (DataRow tableRow in table.Rows)
            {
                var item = Activator.CreateInstance<T>();
                for (var i = 0; i < props.Length; i++)
                {
                    props[i].SetValue(item,
                        tableRow.ItemArray[i].GetType() != typeof(DBNull)
                            ? Convert.ChangeType(tableRow.ItemArray[i], Nullable.GetUnderlyingType(props[i].PropertyType) ?? props[i].PropertyType, CultureInfo.InvariantCulture)
                            : null);
                }

                result.Add(item);
            }

            return result;
        }
    }
}