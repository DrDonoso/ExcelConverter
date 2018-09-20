using System;

namespace ExcelConverter.Extensions
{
    public static class ObjectExtensions
    {
        public static bool IsNull(this object input)
        {
            switch (input)
            {
                case DBNull _:
                    return true;
                case string _:
                    return string.IsNullOrEmpty((string)input);
                default:
                    return false;
            }
        }
    }
}
