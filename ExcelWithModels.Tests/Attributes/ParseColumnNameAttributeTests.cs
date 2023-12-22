namespace ExcelWithModels.Attributes
{
    [TestClass]
    public class ParseColumnNameAttributeTests
    {
        public class TestModel
        {
            [ExcelColumnName("Full Name")]
            public string? Name { get; set; }
        }

        [TestMethod]
        public void CustomColumnName()
        {
            // Arrange
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Full Name";    // Headers   
            worksheet.Cells[2, 1].Value = "John Smith"; // Columns

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            Assert.AreEqual(1, models.Count);
            var model = models.First();
            Assert.AreEqual("John Smith", model.Name);
        }
    }
}
