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
            if (type.Equals(typeof(string)) || type.Equals(typeof(DateTime)))
            {
                return CellValues.String;
            }
            else if (type.Equals(typeof(bool)))
            {
                return CellValues.Boolean;
            }
            // 日付型で書き込むとExcelが破損判定になってしまう
            // 参考：URL：https://social.msdn.microsoft.com/Forums/vstudio/ja-JP/c171cb83-8819-4d26-b411-e39a1e9ccae9/open-xml-sdk?forum=netfxgeneralja
            //else if (type.Equals(typeof(DateTime)))
            //{
            //    return CellValues.Date;
            //}

            return CellValues.Number;
        }

        /// <summary>
        /// セルから値取得
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static object GetValue(this Cell cell)
        {
            string value = cell.CellValue.InnerText;
            switch (cell.DataType.Value)
            {
                case CellValues.Boolean:
                    return value.Equals("1");
                case CellValues.Date:
                    return DateTime.FromOADate(Convert.ToDouble(value));
                case CellValues.InlineString:
                    return cell.InlineString.Text.InnerText;
                case CellValues.Number:
                    double n;
                    double.TryParse(value, out n);
                    return n;
                case CellValues.String:
                case CellValues.Error:
                default:
                    return value;
            }
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
