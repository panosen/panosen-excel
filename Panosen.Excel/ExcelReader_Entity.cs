using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Panosen.Excel
{
    /// <summary>
    /// ExcelReader
    /// </summary>
    public partial class ExcelReader : ExcelHelper
    {
        //private static readonly DateTime DefaultDay = new DateTime(1900, 1, 1);

        /// <summary>
        /// ReadEntityList
        /// </summary>
        public static List<T> ReadEntityList<T>(string inputFilePath, string tableName) where T : class, new()
        {
            var columnInfoList = GetColumnInfos(typeof(T));

            IWorkbook workbook = new XSSFWorkbook(inputFilePath);
            ISheet workSheet = workbook.GetSheet(tableName);

            var firstRow = workSheet.GetRow(0);
            SetColumnIndex(columnInfoList, firstRow);

            List<T> entityList = new List<T>();
            for (int i = 1; i <= workSheet.LastRowNum; i++)
            {
                IRow row = workSheet.GetRow(i);

                var entity = new T();
                entityList.Add(entity);

                foreach (var columnInfo in columnInfoList)
                {
                    if (columnInfo.ColumnIndex == null)
                    {
                        continue;
                    }

                    var cell = row.GetCell(columnInfo.ColumnIndex.Value);

                    SetPropertyValue(entity, columnInfo, cell);
                }
            }

            return entityList;
        }

        private static void SetColumnIndex(List<ColumnInfo> columnInfoList, IRow row)
        {
            for (int i = 0; i < row.LastCellNum; i++)
            {
                foreach (var columnInfo in columnInfoList)
                {
                    if (columnInfo.ColumnAttribute.ColumnName.Equals(row.GetCell(i, MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue))
                    {
                        columnInfo.ColumnIndex = i;
                        break;
                    }
                }
            }
        }

        private static void SetPropertyValue<T>(T entity, ColumnInfo columnInfo, ICell cell) where T : class, new()
        {
            if (cell == null)
            {
                return;
            }

            if (cell.CellType == CellType.Blank)
            {
                return;
            }

            PropertyInfo propertyInfo = columnInfo.PropertyInfo;

            switch (cell.CellType)
            {
                case CellType.Numeric:
                    {
                        switch (propertyInfo.PropertyType.ToString())
                        {
                            case "System.Int32":
                                {
                                    propertyInfo.SetValue(entity, Convert.ToInt32(cell.NumericCellValue), null);
                                }
                                break;
                            case "System.Int64":
                                {
                                    propertyInfo.SetValue(entity, Convert.ToInt64(cell.NumericCellValue), null);
                                }
                                break;
                            case "System.Single":
                                {
                                    propertyInfo.SetValue(entity, Convert.ToSingle(cell.NumericCellValue), null);
                                }
                                break;
                            case "System.Double":
                                {
                                    propertyInfo.SetValue(entity, Convert.ToDouble(cell.NumericCellValue), null);
                                }
                                break;
                            case "System.Decimal":
                                {
                                    propertyInfo.SetValue(entity, Convert.ToDecimal(cell.NumericCellValue), null);
                                }
                                break;
                            default:
                                break;
                        }
                    }
                    break;
                case CellType.String:
                    {
                        switch (propertyInfo.PropertyType.ToString())
                        {
                            case "System.Int32":
                                {
                                    propertyInfo.SetValue(entity, Convert.ToInt32(cell.StringCellValue), null);
                                }
                                break;
                            case "System.Int64":
                                {
                                    propertyInfo.SetValue(entity, Convert.ToInt64(cell.StringCellValue), null);
                                }
                                break;
                            case "System.String":
                                {
                                    propertyInfo.SetValue(entity, cell.StringCellValue, null);
                                }
                                break;
                            default:
                                break;
                        }
                    }
                    break;
                case CellType.Boolean:
                    {
                        switch (propertyInfo.PropertyType.ToString())
                        {
                            case "System.String":
                                {
                                    if (cell.BooleanCellValue)
                                    {
                                        propertyInfo.SetValue(entity, "是", null);
                                    }
                                    else
                                    {
                                        propertyInfo.SetValue(entity, "否", null);
                                    }
                                }
                                break;
                            case "System.Boolean":
                                {
                                    propertyInfo.SetValue(entity, cell.BooleanCellValue, null);
                                }
                                break;
                        }
                    }
                    break;
                case CellType.Error:
                case CellType.Formula:
                case CellType.Blank:
                case CellType.Unknown:
                default:
                    break;
            }
        }
    }
}
