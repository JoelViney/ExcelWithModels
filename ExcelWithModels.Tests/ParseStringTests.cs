namespace ExcelWithModels
{
    [TestClass]
    public class ParseStringTests
    {
        public class TestModel
        {
            public string? Name { get; set; }
        }

        [TestMethod]
        public void ParseString()
        {
            // Arrange
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Name";       // Headers   
            worksheet.Cells[2, 1].Value = "John Smith"; // Columns

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();
            Assert.AreEqual("John Smith", model?.Name);
        }

        [TestMethod]
        public void ParseStringNullIsEmpty()
        {
            // Arrange
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Name";   // Headers   
            worksheet.Cells[2, 1].Value = null;     // Columns

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            var model = models.FirstOrDefault();
            Assert.AreEqual("", model?.Name); 
        }
    }
}