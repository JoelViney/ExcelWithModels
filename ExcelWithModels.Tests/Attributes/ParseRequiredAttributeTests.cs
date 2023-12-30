namespace ExcelWithModels.Attributes
{
    [TestClass]
    public class ParseRequiredAttributeTests
    {
        public class TestModel
        {
            [ExcelRequired]
            public string? Name { get; set; }
            public string? Note { get; set; }
        }

        [TestMethod]
        public void ParseRequiredIsEmptyString()
        {
            // Arrange
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Name";   // Headers   
            worksheet.Cells[1, 2].Value = "Note";
            worksheet.Cells[2, 1].Value = null;       // Columns
            worksheet.Cells[2, 2].Value = "This is a note";

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();
            Assert.AreEqual("", model?.Name); // Strings don't really support null in excel.

            Assert.AreEqual(1, validations.Count);
            var validation = validations.First();
            Assert.AreEqual(2, validation.Row);
            Assert.AreEqual("The string field 'Name' is a required field.", validation.Message);
        }

        [TestMethod]
        public void ParseRequiredIsNullString()
        {
            // Arrange
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Name";   // Headers   
            worksheet.Cells[1, 2].Value = "Note";
            worksheet.Cells[2, 1].Value = "";       // Columns
            worksheet.Cells[2, 2].Value = "This is a note";

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();
            Assert.AreEqual("", model?.Name);

            Assert.AreEqual(1, validations.Count);
            var validation = validations.First();
            Assert.AreEqual(2, validation.Row);
            Assert.AreEqual("The string field 'Name' is a required field.", validation.Message);
        }
    }
}
