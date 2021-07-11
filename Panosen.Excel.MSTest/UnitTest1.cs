using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;

namespace Panosen.Excel.MSTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            //var bookEntityList = ExcelReader.ReadEntityList<BookEntity>(@"F:\book.xlsx", "Õº È");

            var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("D"));
            var tableName = "Õº È";

            List<BookEntity> bookEntityList = new List<BookEntity>();
            for (int i = 0; i < 3; i++)
            {
                BookEntity bookEntity = new BookEntity();
                bookEntity.Id = i;
                bookEntity.Name = $"Book{i}";
                bookEntityList.Add(bookEntity);
            }

            ExcelWriter.WriteEntityList(path, tableName, bookEntityList);

            var actual = ExcelReader.ReadEntityList<BookEntity>(path, tableName);

            Assert.AreEqual(3, actual.Count);
            for (int i = 0; i < 3; i++)
            {
                Assert.AreEqual(i, actual[i].Id);
                Assert.AreEqual($"Book{i}", actual[i].Name);
            }

        }
    }
}
