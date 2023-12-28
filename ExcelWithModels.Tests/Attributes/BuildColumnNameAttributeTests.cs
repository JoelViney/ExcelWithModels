namespace ExcelWithModels.Attributes
{
    [TestClass]
    public class BuildColumnNameAttributeTests
    {
        public class TestModel
        {
            [ExcelColumnName("Full Name")]
            public string? Name { get; set; }
        }

        [TestMethod]
        public void BuildTest()
        {
            // Arrange
            var list = new List<TestModel>()
            {
                new TestModel() { Name = "John Smith"}
            };

            using var excel = new ExcelBuilder();

            // Act
            var xls = excel.Build(list);

            // Assert
            var worksheet = xls.Workbook.Worksheets[0];

            Assert.AreEqual("Full Name", worksheet.Cells[1, 1].Value);
            Assert.AreEqual("John Smith", worksheet.Cells[2, 1].Value);
        }
    }
}
