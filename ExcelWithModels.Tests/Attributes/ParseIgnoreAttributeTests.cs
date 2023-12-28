namespace ExcelWithModels.Attributes
{
    [TestClass]
    public class ParseIgnoreAttributeTests
    {
        public class TestModel
        {
            [ExcelIgnore]
            public int Id { get; set; }
            public string? Name { get; set; }
        }

        [TestMethod]
        public void IgnoreProperty()
        {
            // Arrange
            using var excel = new ExcelParser();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Name";    // Headers   
            worksheet.Cells[2, 1].Value = "John Smith"; // Columns

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            Assert.AreEqual(1, models.Count);
            var model = models.First();
            Assert.AreEqual("John Smith", model.Name);

            Assert.AreEqual(0, validations.Count);
        }
    }
}
