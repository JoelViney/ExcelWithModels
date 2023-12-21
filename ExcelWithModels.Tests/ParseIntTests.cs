namespace ExcelWithModels
{
    public class TestIntDomainModel
    {
        public int Number { get; set; }
    }

    [TestClass]
    public class ParseIntTests
    {
        [TestMethod]
        public void ParseInt()
        {
            // Arrange
            using var excel = new ExcelModelLibrary();

            var worksheet = excel.CreateWorksheet();
            worksheet.Cells[1, 1].Value = "Number"; // Headers   
            worksheet.Cells[2, 1].Value = 23;       // Columns

            // Act
            var (models, validations) = excel.Parse<TestIntDomainModel>(worksheet);

            // Assert
            Assert.AreEqual(1, models.Count);
            var model = models.First();
            Assert.AreEqual(23, model.Number);
        }

    }
}
