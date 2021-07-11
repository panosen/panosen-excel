using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Panosen.Excel
{
    /// <summary>
    /// 单元格格式
    /// </summary>
    public enum ExcelCellFormat
    {
        /// <summary>
        /// 默认
        /// </summary>
        Default = 0,

        /// <summary>
        /// 文本
        /// </summary>
        Text = 1,

        /// <summary>
        /// 保留两位小数
        /// </summary>
        Number2 = 2
    }
}
