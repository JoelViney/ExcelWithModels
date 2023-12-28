namespace ExcelWithModels.Attributes
{
    [TestClass]
    public class BuildIgnoreAttributeTests
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
            var list = new List<TestModel>()
            {
                new TestModel() { Id = 1, Name = "John Smith"}
            };

            using var excel = new ExcelBuilder();

            // Act
            var xls = excel.Build(list);

            // Assert
            var worksheet = xls.Workbook.Worksheets[0];

            Assert.AreEqual("Name", worksheet.Cells[1, 1].Value);
            Assert.AreEqual("John Smith", worksheet.Cells[2, 1].Value);
        }
    }
}
