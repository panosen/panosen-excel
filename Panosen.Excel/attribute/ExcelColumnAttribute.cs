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
        /// ExcelColumnAttribute
        /// </summary>
        public ExcelColumnAttribute(string columnName)
        {
            this.ColumnName = columnName;
        }

        /// <summary>
        /// 单元格格式
        /// </summary>
        public ExcelCellFormat ExcelCellFormat { get; set; }
    }
}
