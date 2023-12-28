using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWithModels
{
    [TestClass]
    public class ParseNullableIntTests
    {
        public class TestNullableIntDomainModel
        {
            public int? Number { get; set; }
            public string? Note { get; set; }
        }

        [TestMethod]
        public void ParseNullableInt()
        {
            // Arrange
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Number"; // Headers
            worksheet.Cells[1, 2].Value = "Note";
            worksheet.Cells[2, 1].Value = 23;       // Columns

            // Act
            var (models, validations) = excel.Parse<TestNullableIntDomainModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();
            Assert.AreEqual(23, model?.Number);
        }

        [TestMethod]
        public void ParseNullableIntAsNull()
        {
            // Arrange
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Number"; // Headers
            worksheet.Cells[1, 2].Value = "Note";
            worksheet.Cells[2, 1].Value = null;     // Columns
            worksheet.Cells[2, 2].Value = "This is a note";

            // Act
            var (models, validations) = excel.Parse<TestNullableIntDomainModel>(worksheet);

            // Assert
            Assert.AreEqual(1, models.Count);
            var model = models.First();
            Assert.IsNull(model?.Number);
        }
    }
}
