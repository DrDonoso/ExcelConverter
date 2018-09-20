using ExcelDataReader;
using System.Data;
using System.IO;

namespace ExcelConverter.Extensions
{
    public static class StreamExtensions
    {
        public static DataTable ToDataTable(this Stream input)
        {
            var reader = ExcelReaderFactory.CreateOpenXmlReader(input);
            var dataSet = reader.AsDataSet();

            var table = dataSet.Tables[0];

            return table;
        }
    }
}
