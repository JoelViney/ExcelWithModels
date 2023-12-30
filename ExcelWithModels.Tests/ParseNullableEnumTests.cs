using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWithModels
{
    [TestClass]
    public class ParseNullableEnumTests
    {
        public enum TestEnum
        {
            NotSpecified = 0,
            Green = 1,
            Blue = 2,
            Red = 4
        }

        public class TestModel
        {
            public TestEnum? Colour { get; set; }
            public string? Note { get; set; }
        }

        [TestMethod]
        public void ParseNullableEnum()
        {
            // Arrange
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Colour"; // Headers
            worksheet.Cells[1, 2].Value = "Note";
            worksheet.Cells[2, 1].Value = "Green";  // Columns

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();
            Assert.AreEqual(TestEnum.Green, model?.Colour);
        }

        [TestMethod]
        public void ParseNullableEnumIsNullReturnsNull()
        {
            // Arrange
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Colour"; // Headers
            worksheet.Cells[1, 2].Value = "Note";
            worksheet.Cells[2, 1].Value = null;     // Columns
            worksheet.Cells[2, 2].Value = "This is a note";

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            Assert.AreEqual(1, models.Count);
            var model = models.First();
            Assert.IsNull(model?.Colour);
        }
    }
}
