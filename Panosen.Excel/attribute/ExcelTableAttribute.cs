using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Panosen.Excel
{
    /// <summary>
    /// 对应一张Excel表格
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public class ExcelTableAttribute : Attribute
    {
    }
}
