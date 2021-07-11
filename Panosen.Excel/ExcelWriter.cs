using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Panosen.Excel
{
    /// <summary>
    /// ExcelWriter
    /// </summary>
    public class ExcelWriter : ExcelHelper
    {
        /// <summary>
        /// 把实体写入Excel
        /// </summary>
        public static void WriteEntityList<T>(string filePath, string tableName, List<T> entityList) where T : class, new()
        {
            var type = typeof(T);

            var attributes = (ExcelTableAttribute[])type.GetCustomAttributes(typeof(ExcelTableAttribute), false);
            if (attributes.Length == 0)
            {
                throw new Exception($"ExcelTableAttribute must be defined on type {type.FullName}");
            }

            IWorkbook workbook = new XSSFWorkbook();

            ISheet workSheet = workbook.CreateSheet(tableName);

            var columnInfoList = GetColumnInfos(typeof(T));

            var rowIndex = 0;

            WriteHeader(workSheet.CreateRow(rowIndex++), columnInfoList);

            foreach (var entity in entityList)
            {
                var row = workSheet.CreateRow(rowIndex++);
                var columnIndex = 0;
                foreach (var columnInfo in columnInfoList)
                {
                    var cell = row.CreateCell(columnIndex++);

                    var value = columnInfo.PropertyInfo.GetValue(entity, null);
                    if (value == null)
                    {
                        continue;
                    }

                    WriteCell(cell, columnInfo, value);
                }
            }

            workbook.Write(new FileStream(filePath, FileMode.CreateNew));
        }

        private static void WriteHeader(IRow headerRow, List<ColumnInfo> columnInfoList)
        {
            int columnIndex = 0;

            foreach (var columnInfo in columnInfoList)
            {
                headerRow.CreateCell(columnIndex++).SetCellValue(columnInfo.ColumnAttribute.ColumnName);
            }
        }

        private static void WriteCell(ICell cell, ColumnInfo columnInfo, object value)
        {
            PropertyInfo propertyInfo = columnInfo.PropertyInfo;

            switch (propertyInfo.PropertyType.ToString())
            {
                case "System.Int32":
                case "System.Int64":
                case "System.Single":
                case "System.Double":
                case "System.Decimal":
                    {
                        var doubleValue = Convert.ToDouble(value);
                        cell.SetCellValue(doubleValue);
                    }
                    break;
                case "System.DateTime":
                    {
                        cell.SetCellValue((DateTime)value);
                    }
                    break;
                case "System.Boolean":
                    {
                        cell.SetCellValue((bool)value);
                    }
                    break;
                case "System.String":
                    {
                        cell.SetCellValue((string)value);
                    }
                    break;
                default:
                    break;
            }
        }
    }
}
