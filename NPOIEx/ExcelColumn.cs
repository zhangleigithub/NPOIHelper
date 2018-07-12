using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOIEx
{
    [AttributeUsage(AttributeTargets.All)]
    public class ExcelColumn : Attribute
    {
        /// <summary>
        /// 标签名称
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="columnIndex">列索引</param>
        public ExcelColumn(int columnIndex)
        {
            ColumnIndex = columnIndex;
        }
    }
}
