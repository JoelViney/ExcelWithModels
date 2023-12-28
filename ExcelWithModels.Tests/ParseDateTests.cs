namespace ExcelWithModels
{
    [TestClass]
    public class ParseDateTests
    {
        public class TestModel
        {
            public DateTime Date { get; set; }
            public string? Note { get; set; }
        }

        [TestMethod]
        public void ParseDate()
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
        public void ParseDateValue()
        {
            // Arrange
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Date";       // Headers
            worksheet.Cells[1, 2].Value = "Note";
            worksheet.Cells[2, 1].Value = new DateTime(2023, 09, 12); // Columns

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();
            Assert.AreEqual(new DateTime(2023, 09, 12), model?.Date);
        }

        [TestMethod]
        public void ParseDateIsNullReturnsValidation()
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
            var model = models.FirstOrDefault();
            Assert.AreEqual(new DateTime(), model?.Date);

            Assert.AreEqual(1, validations.Count);
            var validation = validations.First();
            Assert.AreEqual(2, validation.Row);
            Assert.AreEqual("The date field 'Date' was not populated.", validation.Message);
        }
    }
}
