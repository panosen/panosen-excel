using System;
using System.Collections.Generic;
using System.Text;

namespace Panosen.Excel.MSTest
{
    [ExcelTable()]
    public class BookEntity
    {
        [ExcelColumn("序号")]
        public int Id { get; set; }

        [ExcelColumn("书名")]
        public string Name { get; set; }
    }
}
