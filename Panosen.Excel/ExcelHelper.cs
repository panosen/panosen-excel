using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Panosen.Excel
{
    /// <summary>
    /// ExcelHelper
    /// </summary>
    public class ExcelHelper
    {
        /// <summary>
        /// GetColumnInfos
        /// </summary>
        protected static List<ColumnInfo> GetColumnInfos(Type type)
        {
            List<ColumnInfo> columnInfoList = new List<ColumnInfo>();

            var properties = type.GetProperties();
            foreach (var property in properties)
            {
                if (!property.IsDefined(typeof(ExcelColumnAttribute), false))
                {
                    continue;
                }

                var columnAttribute = (ExcelColumnAttribute)property.GetCustomAttributes(typeof(ExcelColumnAttribute), false)[0];

                columnInfoList.Add(new ColumnInfo { PropertyInfo = property, ColumnAttribute = columnAttribute });
            }

            return columnInfoList;
        }
    }
}
