using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Panosen.Excel
{
    /// <summary>
    /// Excel列
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelColumnAttribute : Attribute
    {
        /// <summary>
        /// Excel列名
        /// </summary>
        public string ColumnName { get; private set; }

        /// <summary>
        /// 列顺序，从1开始
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// ExcelColumnAttribute
        /// </summary>
        public ExcelColumnAttribute(string columnName)
        {
            this.ColumnName = columnName;
        }

        /// <summary>
        /// ExcelColumnAttribute
        /// </summary>
        /// <param name="columnIndex"></param>
        public ExcelColumnAttribute(int columnIndex)
        {
            this.ColumnIndex = columnIndex;
        }

        /// <summary>
        /// ExcelColumnAttribute
        /// </summary>
        public ExcelColumnAttribute(string columnName, int columnIndex)
        {
            this.ColumnName = columnName;
            this.ColumnIndex = columnIndex;
        }

        /// <summary>
        /// 单元格格式
        /// </summary>
        public ExcelCellFormat ExcelCellFormat { get; set; }
    }
}
