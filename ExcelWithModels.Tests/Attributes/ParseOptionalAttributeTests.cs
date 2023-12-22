namespace ExcelWithModels.Attributes
{
    [TestClass]
    public class ParseOptionalAttributeTests
    {
        public class TestModel
        {
            public string? Name { get; set; }

            [ExcelOptional]
            public DateTime DateOfBirth {get;set;}
        }

        [TestMethod]
        public void MissingOptionalColumn()
        {
            // Arrange
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Name";    // Headers
            worksheet.Cells[2, 1].Value = "John Smith"; // Columns

            // Act
            var (models, validations) = excel.Parse<TestModel>(worksheet);

            // Assert
            var model = models.First();
            Assert.AreEqual("John Smith", model.Name);

            Assert.AreEqual(0, validations.Count);
        }
    }
}
