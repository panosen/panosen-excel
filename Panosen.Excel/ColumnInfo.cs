using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Panosen.Excel
{
    /// <summary>
    /// ColumnInfo
    /// </summary>
    public class ColumnInfo
    {
        /// <summary>
        /// PropertyInfo
        /// </summary>
        public PropertyInfo PropertyInfo { get; set; }

        /// <summary>
        /// ColumnAttribute
        /// </summary>
        public ExcelColumnAttribute ColumnAttribute { get; set; }
    }
}
