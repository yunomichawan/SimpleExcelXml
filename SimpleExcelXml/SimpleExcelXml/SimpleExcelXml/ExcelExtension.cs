using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

namespace SimpleExcelXml
{
    public static class ExcelExtension
    {
        /// <summary>
        /// セルの書込型を取得
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static CellValues GetCellValuesType(this Type type)
        {
            if (type.Equals(typeof(string)))
            {
                return CellValues.String;
            }
            else if (type.Equals(typeof(bool)))
            {
                return CellValues.Boolean;
            }
            else if (type.Equals(typeof(DateTime)))
            {
                return CellValues.Date;
            }

            return CellValues.Number;
        }
    }

    public static class CommonExtension
    {
        public static void ForEach<T>(this IEnumerable<T> enumerable, Action<T> callback)
        {
            foreach (T obj in enumerable)
            {
                callback(obj);
            }
        }
    }

}
