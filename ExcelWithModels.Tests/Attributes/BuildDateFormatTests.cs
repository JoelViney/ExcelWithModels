namespace ExcelWithModels.Attributes
{
    [TestClass]
    public class BuildDateFormatTests
    {
        public class TestModel
        {
            [ExcelFormat("dd/MM/yyyy")]
            public DateTime Date { get; set; }
        }

        [TestMethod]
        public void BuildDateFormatTest()
        {
            // Arrange
            var list = new List<TestModel>()
            {
                new TestModel() { Date = new DateTime(2023, 02, 11)}
            };

            using var excel = new ExcelBuilder();

            // Act
            var xls = excel.Build(list);

            // Assert
            var worksheet = xls.Workbook.Worksheets[0];

            Assert.AreEqual("Date", worksheet.Cells[1, 1].Value);
            Assert.AreEqual("11/02/2023", worksheet.Cells[2, 1].Value);
        }
    }
}
