namespace ExcelWithModels
{
    [TestClass]
    public class ParseEnumTests
    {
        public enum TestEnum
        {
            NotSpecified = 0,
            Green = 1,
            Blue = 2,
            RedOrOrange = 4
        }

        public class TestModel
        {
            public TestEnum Colour { get; set; }
            public string? Note { get; set; }
        }

        [TestMethod]
        public void ParseEnum()
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
        public void ParseEnumValue()
        {
            // Arrange
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Colour"; // Headers
            worksheet.Cells[1, 2].Value = "Note";
            worksheet.Cells[2, 1].Value = "1";      // Columns

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();
            Assert.AreEqual(TestEnum.Green, model?.Colour);
        }

        [TestMethod]
        public void ParseEnumIsNullReturnsValidation()
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
            var model = models.FirstOrDefault();
            Assert.AreEqual(TestEnum.NotSpecified, model?.Colour);

            Assert.AreEqual(1, validations.Count);
            var validation = validations.First();
            Assert.AreEqual(2, validation.Row);
            Assert.AreEqual("The enum field 'Colour' was not populated.", validation.Message);
        }

        [TestMethod]
        public void ParseWordifiedNameEnum()
        {
            // Arrange
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Colour"; // Headers
            worksheet.Cells[1, 2].Value = "Note";
            worksheet.Cells[2, 1].Value = "Red Or Orange"; // Columns

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();
            Assert.AreEqual(TestEnum.RedOrOrange, model?.Colour);
        }
    }
}
