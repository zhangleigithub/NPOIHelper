using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace NPOIEx
{
    public static class ExcelEx
    {
        /// <summary>
        /// 扩展方法
        /// </summary>
        /// <param name="worksheet">sheet</param>
        /// <returns>泛型集合</returns>
        public static IEnumerable<T> ConvertSheetToObjects<T>(this ISheet worksheet) where T : new()
        {
            Func<object, bool> func = y => y.GetType() == typeof(ExcelColumn);

            var columns = typeof(T).GetProperties().Where(x => x.GetCustomAttributes(true).Any(func)).Select(p => new
            {
                Property = p,
                ColumnIndex = ((ExcelColumn)p.GetCustomAttributes(true).First(x => x.GetType() == typeof(ExcelColumn))).ColumnIndex
            }).ToList();

            List<T> result = new List<T>();

            for (int i = worksheet.FirstRowNum; i <= worksheet.LastRowNum; i++)
            {
                var tnew = new T();

                columns.ForEach(col =>
                {
                    ICell cell = worksheet.GetRow(i).Cells[col.ColumnIndex];

                    object value = null;

                    if (cell.CellType == CellType.String)
                    {
                        value = cell.StringCellValue;
                    }
                    else if (cell.CellType == CellType.Numeric)
                    {
                        value = cell.NumericCellValue;
                    }
                    else if (cell.CellType == CellType.Boolean)
                    {
                        value = cell.BooleanCellValue;
                    }

                    if (col.Property.PropertyType == typeof(int))
                    {
                        col.Property.SetValue(tnew, Convert.ToInt32(value), null);
                    }
                    else if (col.Property.PropertyType == typeof(string))
                    {
                        col.Property.SetValue(tnew, Convert.ToString(value), null);
                    }
                    else if (col.Property.PropertyType == typeof(bool))
                    {
                        col.Property.SetValue(tnew, Convert.ToBoolean(value), null);
                    }

                });

                result.Add(tnew);
            }

            return result;
        }
    }
}
