using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Panosen.Excel
{
    partial class ExcelReader
    {

        /// <summary>
        /// ReadMapList
        /// </summary>
        public static List<Dictionary<string, string>> ReadMapList(string inputFilePath, string tableName)
        {
            IWorkbook workbook = new XSSFWorkbook(inputFilePath);
            ISheet workSheet = workbook.GetSheet(tableName);

            var firstRow = workSheet.GetRow(0);
            var columnNames = GetColumnNames(firstRow);

            List<Dictionary<string, string>> entityList = new List<Dictionary<string, string>>();
            for (int i = 1; i <= workSheet.LastRowNum; i++)
            {
                IRow row = workSheet.GetRow(i);

                var map = new Dictionary<string, string>();
                entityList.Add(map);

                for (int columnIndex = 0; columnIndex < columnNames.Count; columnIndex++)
                {
                    var columnName = columnNames[columnIndex];
                    if (string.IsNullOrEmpty(columnName))
                    {
                        continue;
                    }

                    var cell = row.GetCell(columnIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    if (cell == null)
                    {
                        continue;
                    }

                    var cellValue = GetCellValueAsString(cell);

                    map[columnName] = cellValue;
                }
            }

            return entityList;
        }

        private static List<string> GetColumnNames(IRow row)
        {
            List<string> columnNames = new List<string>();

            for (int i = 0; i < row.LastCellNum; i++)
            {
                columnNames.Add(row.GetCell(i, MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue);
            }

            return columnNames;
        }

        private static string GetCellValueAsString(ICell cell)
        {
            if (cell.CellType == CellType.Blank)
            {
                return string.Empty;
            }

            switch (cell.CellType)
            {
                case CellType.Numeric:
                    {
                        return cell.NumericCellValue.ToString();
                    }

                case CellType.String:
                    {
                        return cell.StringCellValue;
                    }

                case CellType.Boolean:
                    {
                        return cell.BooleanCellValue.ToString();
                    }

                case CellType.Error:
                case CellType.Formula:
                case CellType.Blank:
                case CellType.Unknown:
                default:
                    {
                        return string.Empty;
                    }
            }
        }
    }
}
