﻿namespace ExcelWithModels
{
    [TestClass]
    public class ParseNullableDateTests
    {
        public class TestModel
        {
            public DateTime? Date { get; set; }
            public string? Note { get; set; }
        }

        [TestMethod]
        public void ParseNullableDate()
        {
            // Arrange
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Date";       // Headers
            worksheet.Cells[1, 2].Value = "Note";
            worksheet.Cells[2, 1].Value = "2023-09-12"; // Columns

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();
            Assert.AreEqual(new DateTime(2023, 09, 12), model?.Date);
        }

        [TestMethod]
        public void ParseNullableDateIsNullReturnsNull()
        {
            // Arrange
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Date";   // Headers
            worksheet.Cells[1, 2].Value = "Note";
            worksheet.Cells[2, 1].Value = null;     // Columns
            worksheet.Cells[2, 2].Value = "This is a note";

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            Assert.AreEqual(1, models.Count);
            var model = models.First();
            Assert.IsNull(model?.Date);
        }
    }
}
