namespace ExcelWithModels
{
    public class TestStringDomainModel 
    {
        public string? Name { get; set; }
    }

    [TestClass]
    public class ParseStringTests
    {
        [TestMethod]
        public void ParseString()
        {
            // Arrange
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Name";       // Headers   
            worksheet.Cells[2, 1].Value = "John Smith"; // Columns

            // Act
            var (models, validations) = excel.Parse<TestStringDomainModel>(worksheet);

            // Assert
            Assert.AreEqual(1, models.Count);
            var model = models.First();
            Assert.AreEqual("John Smith", model.Name);
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
            var (models, validations) = excel.Parse<TestStringDomainModel>(worksheet);

            // Assert
            Assert.AreEqual(1, models.Count);
            var model = models.First();
            Assert.AreEqual("", model.Name); 
        }
    }
}