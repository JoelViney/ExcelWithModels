namespace ExcelWithModels
{
    [TestClass]
    public class ParseBooleanTests
    {
        public class TestModel
        {
            public bool TrueOrFalse { get; set; }
            public string? Note { get; set; }
        }

        [TestMethod]
        public void ParseBoolean()
        {
            // Arrange
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "TrueOrFalse"; // Headers
            worksheet.Cells[1, 2].Value = "Note";
            worksheet.Cells[2, 1].Value = true;       // Columns

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();
            Assert.AreEqual(true, model?.TrueOrFalse);
        }

        [TestMethod]
        public void ParseBooleanString()
        {
            // Arrange
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "TrueOrFalse"; // Headers
            worksheet.Cells[1, 2].Value = "Note";
            worksheet.Cells[2, 1].Value = "true";       // Columns

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();
            Assert.AreEqual(true, model?.TrueOrFalse);
        }

        [TestMethod]
        public void ParseBooleanIsNullReturnsValidation()
        {
            // Arrange
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "TrueOrFalse"; // Headers
            worksheet.Cells[1, 2].Value = "Note";
            worksheet.Cells[2, 1].Value = null;         // Columns
            worksheet.Cells[2, 2].Value = "This is a note";

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();
            Assert.AreEqual(false, model?.TrueOrFalse); // default value

            Assert.AreEqual(1, validations.Count);
            var validation = validations.First();
            Assert.AreEqual(2, validation.Row);
            Assert.AreEqual("The boolean field 'TrueOrFalse' was not populated.", validation.Message);
        }
    }
}
